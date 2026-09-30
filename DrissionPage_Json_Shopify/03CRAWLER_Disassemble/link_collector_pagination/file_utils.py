# ======================== 文件持久化层 ========================
#
# 职责: 采集结果落盘。含旧文件安全读取、恢复文件命名、分类索引维护、
#       按域名分组合并写 XLSX。
# 依赖: config、logger、openpyxl。
# 禁止: 导入 tkinter、DrissionPage、collector、browser_utils。
#
# 安全约定（在本版中进一步强化）:
#   1. 旧 XLSX 存在但读取失败 -> 返回 None，调用方必须另存 recovery 文件，
#      绝不覆盖原文件。
#   2. XLSX / JSON 均采用“临时文件 + os.replace”的原子写入，
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
from typing import Dict, Iterable, List, Optional, Sequence
from urllib.parse import urlparse

import openpyxl

import config
import logger


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
    try:
        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)
    except OSError:
        pass


# ======================== 读取旧文件 ========================
def load_existing_domain_rows(file_path: str) -> Optional[Dict[str, str]]:
    """读取已存在的 {domain}.xlsx。

    返回值：
      - 文件不存在：{}，允许直接新建；
      - 读取成功：{link: title}；
      - 文件存在但读取失败：None，调用方必须保护原文件，禁止覆盖。
    """
    existing: Dict[str, str] = {}
    if not os.path.exists(file_path):
        return existing
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            title = str(row[0] or "").strip()
            link = str(row[1] or "").strip()
            if link and link not in existing:
                existing[link] = title
        wb.close()
    except Exception as e:
        logger.log(f"读取已有文件失败，已保护原文件不覆盖: {e}")
        return None
    return existing


def build_recovery_file_path(file_path: str) -> str:
    """为无法安全合并的结果生成不覆盖历史文件的恢复文件路径。

    自动沿用原文件扩展名，因此 JSON 索引会得到 *.recovery-*.json。
    """
    base, ext = os.path.splitext(file_path)
    stamp = time.strftime("%Y%m%d-%H%M%S")
    mark = config.RECOVERY_MARK
    candidate = f"{base}.{mark}-{stamp}{ext or '.xlsx'}"
    counter = 2
    while os.path.exists(candidate):
        candidate = f"{base}.{mark}-{stamp}-{counter}{ext or '.xlsx'}"
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


def _unpack_category(item):
    """兼容 CategoryItem 与 (title, url) 元组两种形式。"""
    title = getattr(item, "title", None)
    url = getattr(item, "url", None)
    if title is None and url is None and isinstance(item, (tuple, list)):
        if len(item) >= 2:
            return item[0], item[1]
        return "", item[0] if item else ""
    return title or "", url or ""


# ======================== 写入 XLSX ========================
def _write_domain_xlsx(file_path: str, rows: Iterable) -> int:
    """按 config.XLSX_HEADERS 写出 {link: title} 映射，返回行数。

    原子写入：先写同目录临时文件，wb.save 成功后再 os.replace。
    中途异常会清理临时文件并抛出，原文件不受影响。
    """
    target_dir = os.path.dirname(file_path) or "."
    os.makedirs(target_dir, exist_ok=True)

    fd, tmp_path = tempfile.mkstemp(prefix=".tmp_xlsx_", suffix=".xlsx", dir=target_dir)
    os.close(fd)
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(list(config.XLSX_HEADERS))
        count = 0
        for link, title in rows:
            ws.append([title, link])
            count += 1
        wb.save(tmp_path)
        wb.close()
        _atomic_replace(tmp_path, file_path)
        return count
    except Exception:
        _cleanup_tmp(tmp_path)
        raise


@dataclass
class SaveReport:
    """一次保存的统计结果，供 GUI 汇总展示。"""

    total_rows: int = 0
    domain_files: int = 0
    written_files: List[str] = field(default_factory=list)
    recovery_files: List[str] = field(default_factory=list)
    json_recovery_files: List[str] = field(default_factory=list)
    index_path: str = ""
    error: Optional[str] = None

    def describe(self) -> str:
        if self.error:
            return f"保存失败: {self.error}"
        if not self.total_rows:
            return "本次没有新增可保存的链接"
        parts = [f"共写入 {self.domain_files} 个域名文件、{self.total_rows} 条链接"]
        if self.recovery_files:
            parts.append(f"其中 {len(self.recovery_files)} 个为恢复文件")
        if self.json_recovery_files:
            parts.append(f"索引备份 {len(self.json_recovery_files)} 个")
        return "；".join(parts)


def save_results(
    rows: Sequence[dict],
    output_path: str,
    categories: Sequence = (),
) -> SaveReport:
    """保存采集结果：更新分类索引 + 按域名分组合并写 XLSX。

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

    # 2) 更新分类索引：失败不影响后续 XLSX 写出
    try:
        report.index_path = update_link_index(
            output_path, categories, recovery_out=report.json_recovery_files
        )
    except Exception as e:
        logger.log(f"更新分类索引失败（不影响 XLSX 写出）: {e}")
        report.error = f"索引更新失败: {e}"

    logger.log(f"采集结束: 共 {len(rows)} 条链接")

    if not rows:
        logger.log("本次没有新增可保存的链接")
        return report

    # 3) 按域名分组
    domain_groups: Dict[str, List[dict]] = {}
    for row in rows:
        domain = urlparse(row["link"]).netloc
        domain_groups.setdefault(domain, []).append(row)

    # 4) 逐域名合并写出，单个域名失败不影响其他域名
    for domain, domain_rows in domain_groups.items():
        safe_domain = _safe_domain_filename(domain)
        file_name = f"{safe_domain}.xlsx"
        file_path = os.path.join(output_path, file_name)

        existing = load_existing_domain_rows(file_path)

        # 4a) 旧文件读取失败：本轮结果另存 recovery，绝不覆盖原文件
        if existing is None:
            recovery_path = build_recovery_file_path(file_path)
            recovery_rows: Dict[str, str] = {}
            for row in domain_rows:
                link = row["link"]
                if link not in recovery_rows:
                    recovery_rows[link] = row.get("title", "")
            try:
                _write_domain_xlsx(recovery_path, recovery_rows.items())
                report.recovery_files.append(recovery_path)
                report.written_files.append(recovery_path)
                logger.log(
                    f"原文件读取失败，已禁止覆盖: {file_path}"
                    f"；本轮 {len(recovery_rows)} 条另存为: {recovery_path}"
                )
            except Exception as e:
                logger.log(f"写入恢复文件也失败: {recovery_path}（{e}）")
            continue

        # 4b) 正常合并
        merged = dict(existing)
        added = 0
        for row in domain_rows:
            link = row["link"]
            if link not in merged:
                merged[link] = row.get("title", "")
                added += 1

        try:
            _write_domain_xlsx(file_path, merged.items())
        except Exception as e:
            # 原子写失败：原文件未被破坏，把合并结果另存 recovery
            recovery_path = build_recovery_file_path(file_path)
            try:
                _write_domain_xlsx(recovery_path, merged.items())
                report.recovery_files.append(recovery_path)
                report.written_files.append(recovery_path)
                logger.log(
                    f"写入失败，原文件已保护: {file_path}；"
                    f"合并结果另存为: {recovery_path}（{e}）"
                )
            except Exception as e2:
                logger.log(f"写入恢复文件也失败: {recovery_path}（{e2}）")
            continue

        report.written_files.append(file_path)
        logger.log(
            f"已保存: {file_path}（已有 {len(existing)} 条 + 新增 {added} 条"
            f" = 合并 {len(merged)} 条）"
        )

    report.domain_files = len(domain_groups)
    logger.log(f"共 {len(domain_groups)} 个域名文件")
    return report