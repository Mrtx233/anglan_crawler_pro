# ======================== 文件持久化层 ========================
#
# 职责: 采集结果落盘。含旧文件安全读取、恢复文件命名、分类索引维护、
#       按域名分组合并写 JSON。
# 依赖: config、logger。
# 禁止: 导入 tkinter、DrissionPage、collector、browser_utils。
#
# 输出约定:
#   每个域名一个 {域名}.json，内容为 分类名 -> 商品链接数组，
#   例如 {"Prom Dresses": ["https://x.com/products/a", ...]}。
#   第二、三阶段直接读这些文件，不再经过 Excel。
#
# 安全约定（在本版中进一步强化）:
#   1. 旧域名 JSON 存在但读取失败 -> 返回 None，调用方必须另存 recovery 文件，
#      绝不覆盖原文件。
#   2. 域名 JSON 与分类索引均采用“临时文件 + os.replace”的原子写入，
#      写入过程中断不会破坏已有文件。
#   3. JSON 索引读取失败时先备份为 recovery，再重建；不再直接覆盖。
#   4. 输出域名文件名经过 Windows 合法化清洗。
# ========================================================

import json
import os
import re
import tempfile
import time
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence
from urllib.parse import urlparse

from stages.link_collection import config
from stages.link_collection import logger


# ======================== 文件名清洗 ========================
# Windows 不允许的字符: < > : " / \ | ? * 以及控制字符
_INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def _safe_domain_filename(domain: str) -> str:
    """把 netloc（可能含端口、IPv6）转成 Windows 合法文件名主干。"""
    safe = _INVALID_FILENAME_CHARS.sub("_", str(domain or ""))
    # Windows 文件名不允许以点或空格结尾
    safe = safe.rstrip(". ")
    return safe or "unknown_domain"


# ======================== 原子写入工具 ========================
def _atomic_replace(tmp_path: str, target_path: str) -> None:
    """把临时文件原子替换为目标文件；同目录调用时同文件系统内原子。"""
    os.makedirs(os.path.dirname(target_path) or ".", exist_ok=True)
    os.replace(tmp_path, target_path)


def _cleanup_tmp(tmp_path: str) -> None:
    """尽力删除写入失败后残留的临时文件；删除失败不抛异常。"""
    try:
        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)
    except OSError:
        pass


# ======================== 读取旧文件 ========================
def load_existing_domain_groups(file_path: str) -> Optional[Dict[str, List[str]]]:
    """读取已存在的 {域名}.json。

    返回值：
      - 文件不存在：{}，允许直接新建；
      - 读取成功：{分类名: [链接, ...]}；
      - 文件存在但读取失败：None，调用方必须保护原文件，禁止覆盖。
    """
    if not os.path.exists(file_path):
        return {}
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            return _normalize_domain_groups(json.load(f))
    except (ValueError, TypeError, OSError) as e:
        logger.log(f"读取已有文件失败，已保护原文件不覆盖: {e}")
        return None


def build_recovery_file_path(file_path: str) -> str:
    """为无法安全合并的结果生成不覆盖历史文件的恢复文件路径。"""
    base, ext = os.path.splitext(file_path)
    stamp = time.strftime("%Y%m%d-%H%M%S")
    mark = config.RECOVERY_MARK
    candidate = f"{base}.{mark}-{stamp}{ext or '.json'}"
    counter = 2
    while os.path.exists(candidate):
        candidate = f"{base}.{mark}-{stamp}-{counter}{ext or '.json'}"
        counter += 1
    return candidate


# ======================== 分类索引 ========================
def _dump_compact_index(index: Dict[str, Dict[str, List[str]]]) -> str:
    """紧凑 JSON：域名/分类保留缩进，URL 数组内联单行，减少换行。"""
    domain_lines = []
    for domain, cats in index.items():
        cat_lines = []
        for title, urls in cats.items():
            urls_str = json.dumps(urls, ensure_ascii=False)  # ["u1","u2"] 内联
            cat_lines.append(
                "    " + json.dumps(str(title), ensure_ascii=False) + ": " + urls_str
            )
        domain_lines.append(
            "  " + json.dumps(str(domain), ensure_ascii=False)
            + ": {\n" + ",\n".join(cat_lines) + "\n  }"
        )
    return "{\n" + ",\n".join(domain_lines) + "\n}"


def _atomic_write_json(json_path: str, index: Dict) -> None:
    """原子写入 JSON 索引：先写同目录临时文件，再 os.replace。"""
    target_dir = os.path.dirname(json_path) or "."
    os.makedirs(target_dir, exist_ok=True)
    fd, tmp_path = tempfile.mkstemp(prefix=".tmp_index_", suffix=".json", dir=target_dir)
    os.close(fd)
    try:
        with open(tmp_path, "w", encoding="utf-8") as f:
            f.write(_dump_compact_index(index))
        _atomic_replace(tmp_path, json_path)
    except Exception:
        _cleanup_tmp(tmp_path)
        raise


def update_link_index(
    output_path: str,
    categories: Sequence,
    recovery_out: Optional[List[str]] = None,
) -> str:
    """把分类列表的 域名→分类标题→分类URL 映射记录进 links_index.json。

    结构: {域名: {分类标题: [分类URL, ...]}}
    已存在的文件会被读入并合并，同一分类下的 URL 按出现顺序追加去重。
    分类标题用列表而非单值：同域名下允许重名分类。

    新增: 读取失败时先把原 JSON 备份为 recovery 文件（不覆盖历史），
    并把备份路径追加到 recovery_out（若提供）。
    """
    json_path = os.path.join(output_path, config.LINK_INDEX_FILENAME)
    index: Dict[str, Dict[str, List[str]]] = {}
    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            index = _normalize_index(loaded)
        except (ValueError, TypeError, OSError) as e:
            backup_path = build_recovery_file_path(json_path)
            try:
                os.replace(json_path, backup_path)
                if recovery_out is not None:
                    recovery_out.append(backup_path)
                logger.log(
                    f"读取已有链接记录失败，原文件已备份为: {backup_path}（{e}）"
                )
            except OSError as move_err:
                logger.log(
                    f"读取已有链接记录失败且备份失败: {e}；备份错误: {move_err}"
                )

    added = 0
    title_counts: Dict[tuple, int] = {}
    for item in categories:
        title, url = _unpack_category(item)
        url = str(url or "").strip()
        if not url:
            continue
        domain = urlparse(url).netloc
        title = str(title or "").strip()
        title_counts[(domain, title)] = title_counts.get((domain, title), 0) + 1

        urls = index.setdefault(domain, {}).setdefault(title, [])
        if url not in urls:
            urls.append(url)
            added += 1

    os.makedirs(output_path, exist_ok=True)
    _atomic_write_json(json_path, index)
    logger.log(f"已更新链接记录: {json_path}（新增 {added} 个分类）")

    repeated = [
        f"{domain} / {title} × {count}"
        for (domain, title), count in title_counts.items()
        if count > 1
    ]
    if repeated:
        logger.log(
            "提示: 以下分类标题在同一域名下重复，索引已保留全部 URL: "
            + "；".join(repeated)
        )
    return json_path


def _normalize_index(loaded) -> Dict[str, Dict[str, List[str]]]:
    """兼容旧版 {域名: {分类标题: 分类URL}}：标量一律升级为单元素列表。"""
    index: Dict[str, Dict[str, List[str]]] = {}
    if not isinstance(loaded, dict):
        return index
    for domain, titles in loaded.items():
        if not isinstance(titles, dict):
            continue
        bucket: Dict[str, List[str]] = {}
        for title, value in titles.items():
            if isinstance(value, dict) or value is None:
                continue
            urls: List[str] = []
            for raw in (value if isinstance(value, list) else [value]):
                raw = str(raw or "").strip()
                if raw:
                    urls.append(raw)
            if urls:
                bucket[str(title)] = urls
        if bucket:
            index[str(domain)] = bucket
    return index


def _normalize_domain_groups(loaded) -> Dict[str, List[str]]:
    """把域名 JSON 收敛成 {分类名: [链接, ...]}。

    兼容手写时把单条链接写成字符串而非数组的情况；分类名为空的行丢弃，
    避免 GUI 里出现无法与设置对上的空分类。
    """
    groups: Dict[str, List[str]] = {}
    if not isinstance(loaded, dict):
        return groups
    for title, value in loaded.items():
        if isinstance(value, dict) or value is None:
            continue
        urls: List[str] = []
        for raw in (value if isinstance(value, list) else [value]):
            raw = str(raw or "").strip()
            if raw and raw not in urls:
                urls.append(raw)
        if urls:
            groups[str(title).strip()] = urls
    return groups


def _unpack_category(item):
    """兼容 CategoryItem 与 (title, url) 元组两种形式。"""
    title = getattr(item, "title", None)
    url = getattr(item, "url", None)
    if title is None and url is None and isinstance(item, (tuple, list)):
        if len(item) >= 2:
            return item[0], item[1]
        return "", item[0] if item else ""
    return title or "", url or ""


# ======================== 写入域名 JSON ========================
def _write_domain_json(file_path: str, groups: Dict[str, List[str]]) -> int:
    """把 {分类名: [链接]} 写成 JSON，返回链接总条数。

    原子写入：先写同目录临时文件，写成功后再 os.replace。
    中途异常会清理临时文件并抛出，原文件不受影响。
    """
    target_dir = os.path.dirname(file_path) or "."
    os.makedirs(target_dir, exist_ok=True)

    fd, tmp_path = tempfile.mkstemp(prefix=".tmp_domain_", suffix=".json", dir=target_dir)
    os.close(fd)
    try:
        with open(tmp_path, "w", encoding="utf-8") as f:
            f.write(_dump_domain_groups(groups))
        _atomic_replace(tmp_path, file_path)
    except Exception:
        _cleanup_tmp(tmp_path)
        raise
    return sum(len(urls) for urls in groups.values())


def _dump_domain_groups(groups: Dict[str, List[str]]) -> str:
    """紧凑 JSON：分类保留缩进，URL 数组内联单行，减少换行。"""
    lines = []
    for title, urls in groups.items():
        urls_str = json.dumps(urls, ensure_ascii=False)  # ["u1","u2"] 内联
        lines.append("  " + json.dumps(str(title), ensure_ascii=False) + ": " + urls_str)
    return "{\n" + ",\n".join(lines) + "\n}"


@dataclass
class SaveReport:
    """一次保存的统计结果，供 GUI 汇总展示。"""

    total_rows: int = 0
    # 实际落盘的链接数：写入失败时会小于 total_rows
    written_rows: int = 0
    domain_files: int = 0
    written_files: List[str] = field(default_factory=list)
    recovery_files: List[str] = field(default_factory=list)
    json_recovery_files: List[str] = field(default_factory=list)
    # 非致命：索引更新失败，但域名 JSON 已正常落盘，不算整体保存失败
    index_error: Optional[str] = None
    error: Optional[str] = None

    def describe(self) -> str:
        """把保存统计拼成一句中文摘要，供日志/GUI 直接展示。

        优先级：整体失败只报错误；无新增链接单独提示；
        否则汇报域名文件数、写入条数、未写入条数、恢复文件与索引备份，
        索引失败作为非致命项追加在末尾。
        """
        if self.error:
            return f"保存失败: {self.error}"
        if not self.total_rows:
            parts = ["本次没有新增可保存的链接"]
        else:
            parts = [
                f"共写入 {self.domain_files} 个域名文件、"
                f"{self.written_rows} 条链接"
            ]
            if self.recovery_files:
                parts.append(f"其中 {len(self.recovery_files)} 个为恢复文件")
            unwritten = self.total_rows - self.written_rows
            if unwritten > 0:
                parts.append(f"另有 {unwritten} 条未能写入")
            if self.json_recovery_files:
                parts.append(f"索引备份 {len(self.json_recovery_files)} 个")
        if self.index_error:
            parts.append(f"索引更新失败（域名文件不受影响）: {self.index_error}")
        return "；".join(parts)


def save_results(
    rows: Sequence[dict],
    output_path: str,
    categories: Sequence = (),
) -> SaveReport:
    """保存采集结果：更新分类索引 + 按域名分组合并写 JSON。

    任何异常都被捕获并记录到 SaveReport.error，不向上抛出。
    """
    report = SaveReport(total_rows=len(rows))

    # 1) 输出目录准备：失败即终止，没有可写位置
    try:
        os.makedirs(output_path, exist_ok=True)
    except OSError as e:
        report.error = f"创建输出目录失败: {e}"
        logger.log(report.error)
        return report

    # 2) 更新分类索引：失败不影响后续域名 JSON 写出
    try:
        # 索引路径由 update_link_index 自己打日志，这里不保留返回值
        update_link_index(
            output_path, categories, recovery_out=report.json_recovery_files
        )
    except Exception as e:
        logger.log(f"更新分类索引失败（不影响域名文件写出）: {e}")
        report.index_error = str(e)

    logger.log(f"采集结束: 共 {len(rows)} 条链接")

    if not rows:
        logger.log("本次没有新增可保存的链接")
        return report

    # 3) 按域名分组，组内再按分类归档
    domain_groups: Dict[str, Dict[str, List[str]]] = {}
    for row in rows:
        link = str(row.get("link") or "").strip()
        if not link:
            continue
        title = str(row.get("title") or "").strip()
        domain_groups.setdefault(urlparse(link).netloc, {}).setdefault(title, []).append(link)

    # 4) 逐域名合并写出，单个域名失败不影响其他域名
    for domain, new_groups in domain_groups.items():
        safe_domain = _safe_domain_filename(domain)
        file_path = os.path.join(output_path, f"{safe_domain}.json")

        existing = load_existing_domain_groups(file_path)

        # 4a) 旧文件读取失败：本轮结果另存 recovery，绝不覆盖原文件
        if existing is None:
            recovery_path = build_recovery_file_path(file_path)
            try:
                report.written_rows += _write_domain_json(recovery_path, new_groups)
                report.recovery_files.append(recovery_path)
                report.written_files.append(recovery_path)
                logger.log(
                    f"原文件读取失败，已禁止覆盖: {file_path}"
                    f"；本轮另存为: {recovery_path}"
                )
            except Exception as e:
                report.error = f"写入恢复文件失败: {recovery_path}（{e}）"
                logger.log(report.error)
            continue

        # 4b) 正常合并
        merged, added = _merge_groups(existing, new_groups)
        try:
            # 写盘失败时这行不会执行，written_rows 保持未计入
            report.written_rows += _write_domain_json(file_path, merged)
        except Exception as e:
            # 原子写失败：原文件未被破坏，把合并结果另存 recovery
            recovery_path = build_recovery_file_path(file_path)
            try:
                report.written_rows += _write_domain_json(recovery_path, merged)
                report.recovery_files.append(recovery_path)
                report.written_files.append(recovery_path)
                logger.log(
                    f"写入失败，原文件已保护: {file_path}；"
                    f"合并结果另存为: {recovery_path}（{e}）"
                )
            except Exception as e2:
                report.error = f"写入恢复文件失败: {recovery_path}（{e2}）"
                logger.log(report.error)
            continue

        report.written_files.append(file_path)
        logger.log(
            f"已保存: {file_path}（新增 {added} 条"
            f"，现共 {sum(len(u) for u in merged.values())} 条）"
        )

    # 以实际写盘成功的文件数为准（含 recovery），而非输入分组数：
    # 某域名正常写入与 recovery 都失败时，输入分组数会高报。
    report.domain_files = len(report.written_files)
    logger.log(
        f"共 {report.domain_files} 个域名文件（输入分组 {len(domain_groups)} 个）"
    )
    return report


def _merge_groups(
    existing: Dict[str, List[str]], new_groups: Dict[str, List[str]]
) -> tuple:
    """把新采集结果并入已有分组，返回 (合并结果, 新增链接数)。

    同一分类下按出现顺序追加去重；已有分类保持原有顺序。
    """
    merged = {title: list(urls) for title, urls in existing.items()}
    added = 0
    for title, urls in new_groups.items():
        bucket = merged.setdefault(title, [])
        for url in urls:
            if url not in bucket:
                bucket.append(url)
                added += 1
    return merged, added


# ======================== 读取域名 JSON（第二/三阶段入口） ========================
def read_domain_groups(file_path: str) -> Dict[str, List[str]]:
    """读取单个 {域名}.json，返回 {分类名: [链接, ...]}。

    格式非法时抛 ValueError，由调用方决定是跳过还是中止——
    这里的失败与「采集时旧文件读不动」性质不同，是输入数据本身有问题。
    """
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            loaded = json.load(f)
    except (ValueError, OSError) as e:
        raise ValueError(f"读取链接文件失败: {file_path}（{e}）") from e
    if not isinstance(loaded, dict):
        raise ValueError(f"链接文件格式错误（应为 JSON 对象）: {file_path}")
    return _normalize_domain_groups(loaded)


def read_targets_from_domain_json(file_path: str) -> List[tuple]:
    """读取单个 {域名}.json，按分类顺序展开为 [(分类名, 链接), ...]。

    与旧 read_targets_from_excel 的输出形状一致，第二阶段的
    group_targets_by_json_url 与 GUI 的分类列表都直接复用。
    """
    targets: List[tuple] = []
    for title, urls in read_domain_groups(file_path).items():
        for url in urls:
            targets.append((title, url))
    return targets