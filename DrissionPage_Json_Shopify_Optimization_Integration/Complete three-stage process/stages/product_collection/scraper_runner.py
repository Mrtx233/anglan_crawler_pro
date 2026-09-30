# ======================== 采集运行器 ========================
#
# 职责: shopify_scraper_workbench 的业务主循环。
#       含文件遍历、续采、浏览器创建、请求重试、解析调用、
#       SAVE_EVERY 自动保存、异常保护保存。
# 依赖: config、logger、browser_utils、network_utils、
#       product_parser、file_utils。
# 禁止: 导入 tkinter —— 本模块必须能脱离 GUI 运行与单测。
#
# 与 GUI 的通信方式:
#   1. stop_event（停止信号）
#   2. on_file_status(index, text, state)（文件级状态）
#   3. on_progress(current, total, text)（进度）
#   4. run() 返回值 all_results
# ============================================================

import os
from typing import Callable, List, Optional

from stages.product_collection import browser_utils
from stages.product_collection import config
from stages.product_collection import file_utils
from stages.product_collection import logger
from stages.product_collection import network_utils
from stages.product_collection import product_parser

# 与原 main.py 保持一致的别名
DELAY = config.DELAY
MAX_RETRIES = config.MAX_RETRIES
RETRY_BASE_DELAY = config.RETRY_BASE_DELAY
SAVE_EVERY = config.SAVE_EVERY

create_browser = browser_utils.create_browser
fetch_exchange_rates = network_utils.fetch_exchange_rates
fetch_json = network_utils.fetch_json
group_targets_by_json_url = network_utils.group_targets_by_json_url
parse_product = product_parser.parse_product
read_targets_from_excel = file_utils.read_targets_from_excel
save_xlsx = file_utils.save_xlsx
load_xlsx = file_utils.load_xlsx
build_output_file_path = file_utils.build_output_file_path
build_recovery_xlsx_path = file_utils.build_recovery_xlsx_path

# 回调签名
FileStatusCallback = Callable[[int, str, str], None]
ProgressCallback = Callable[[float, float, str], None]


class ScraperRunner:
    """采集主循环，与 tkinter 彻底解耦。

    用法:
        runner = ScraperRunner(task_files, task_folder, stop_event,
                               on_file_status=fn1, on_progress=fn2)
        rows = runner.run()   # 阻塞执行，内部完成浏览器创建与保存
    """

    def __init__(
        self,
        task_files,
        task_folder,
        stop_event,
        on_file_status: Optional[FileStatusCallback] = None,
        on_progress: Optional[ProgressCallback] = None,
    ):
        """保存任务文件清单与回调；失败数/恢复文件数在 run() 过程中累计。"""
        self.failed_count = 0
        self.recovery_count = 0
        self.task_files = task_files
        self.task_folder = task_folder
        self.stop_event = stop_event
        self.on_file_status = on_file_status
        self.on_progress = on_progress

    # ---------- 回调分发（单个回调异常不影响主流程） ----------
    def _emit_file_status(self, index, text, state):
        """回调分发（单个回调异常不影响主流程）。"""
        if self.on_file_status is None:
            return
        try:
            self.on_file_status(index, text, state)
        except Exception:
            pass

    def _emit_progress(self, current, total, text=""):
        """回调分发（单个回调异常不影响主流程）。"""
        if self.on_progress is None:
            return
        try:
            self.on_progress(current, total, text)
        except Exception:
            pass

    # ---------- 图片位置解析（与原实现语义一致） ----------
    @staticmethod
    def _parse_positions(text):
        """解析 keep/skip 输入，返回 int 列表或 None。

        - 空字符串 -> None（不过滤）
        - 非空但无有效整数 -> []（保留空集合，会过滤掉全部图片）
        - 解析异常 -> None
        """
        if not text:
            return None
        try:
            result = []
            for value in text.replace("，", ",").split(","):
                value = value.strip().replace("－", "-")
                if value.lstrip("-").isdigit():
                    result.append(int(value))
            return result
        except Exception:
            return None

    # ---------- 图片过滤：文件级默认 + 分类级覆盖 ----------
    @staticmethod
    def _resolve_filter_value(text):
        """把单栏文本解析成 keep/skip 参数。

        "-" 是「显式不过滤」记号，解析成 None 而不是空列表
        （空列表会过滤掉全部图片，那是另一个意思）。
        """
        if str(text or "").strip() == config.IMAGE_FILTER_NONE:
            return None
        return ScraperRunner._parse_positions(text)

    def _resolve_image_filter(self, file_keep, file_skip, category_filter=None):
        """决定某个商品实际使用的 (keep, skip)。

        分类行两栏都留空 -> 整套继承文件行；只要有一栏填了（含 "-"），
        该分类就完全用自己的两栏，留空的那栏表示不过滤、不回退到文件行。

        成对解析而非逐栏继承，是为了避免「分类只填了跳过、却继承了文件的
        保留」这类自相矛盾的结果——保留规则优先于跳过规则，逐栏继承会让
        分类自己填的跳过静默失效。
        """
        category_filter = category_filter or {}
        cat_keep_text = category_filter.get("keep", "")
        cat_skip_text = category_filter.get("skip", "")
        if not str(cat_keep_text).strip() and not str(cat_skip_text).strip():
            return file_keep, file_skip
        return (
            self._resolve_filter_value(cat_keep_text),
            self._resolve_filter_value(cat_skip_text),
        )

    # ---------- 图片过滤：日志渲染 ----------
    @staticmethod
    def _format_position_filter(values):
        """渲染单栏解析结果。None=不过滤；[]=解析不出整数的空列表。"""
        if values is None:
            return "不过滤"
        if not values:
            return "空列表"
        return ",".join(str(value) for value in values)

    @classmethod
    def _format_image_filter(cls, keep, skip):
        """把生效的 (keep, skip) 渲染成日志里能直接看懂的说明。

        只写真正生效的规则：两栏都填时保留优先，跳过会被忽略，这里显式
        标注出来，避免日志写着「跳过 3」但实际没生效。
        """
        if keep is None and skip is None:
            return "不过滤"
        if keep is not None and not keep:
            return "保留 空列表（会过滤掉全部图片）"
        if skip is not None and not skip and keep is None:
            return "跳过 空列表（等于不过滤）"
        keep_text = cls._format_position_filter(keep)
        skip_text = cls._format_position_filter(skip)
        if keep is not None and skip is not None:
            return f"保留 {keep_text}（跳过 {skip_text} 不生效：保留优先）"
        if keep is not None:
            return f"保留 {keep_text}"
        return f"跳过 {skip_text}"

    # ---------- 主流程 ----------
    def run(self) -> List[dict]:
        """执行第二阶段全流程，返回所有已采集行（含续采读入的历史行）。

        先拉汇率，再逐个链接表文件处理：读表去重 → 比对已有输出的
        link-href 得到待采集合 → 启动浏览器逐商品抓 json → 按 SAVE_EVERY
        自动保存 → finally 中无论正常/停止/异常都强制落盘。
        外层 except 还会再做一次兜底保存，随后重新抛出异常。
        """
        current_results = None
        current_output_file = None
        current_file_name = ""
        active_file_index = None

        try:
            fetch_exchange_rates()

            all_results: List[dict] = []
            grand_success = 0
            grand_fail = 0
            file_count = len(self.task_files)

            for fi, file_info in enumerate(self.task_files):
                if self.stop_event.is_set():
                    break

                active_file_index = fi
                self._emit_file_status(fi, "处理中", "running")

                fpath = file_info["path"]
                fname = os.path.basename(fpath)
                current_file_name = fname
                skip_str = file_info["skip"]
                keep_str = file_info.get("keep", "")
                file_keep = self._parse_positions(keep_str)
                file_skip = self._parse_positions(skip_str)
                # {分类名: {"keep": 文本, "skip": 文本}}；缺项表示该文件没有
                # 分类级设置，全部走文件级默认。
                category_filters = file_info.get("categories") or {}

                # 安全输出路径：绝不允许与输入 XLSX 为同一路径。
                output_file = build_output_file_path(fpath)
                os.makedirs(os.path.dirname(output_file), exist_ok=True)
                current_output_file = output_file

                print(f"\n{'=' * 50}")
                print(f"[{fi + 1}/{file_count}] {fname}")
                print(f"  输出: {output_file}")
                print(
                    "  图片设置 · 文件级（分类未单独设置时生效）: "
                    + self._format_image_filter(file_keep, file_skip)
                )
                for cat_title, cat_cfg in category_filters.items():
                    if not (
                        str(cat_cfg.get("keep", "")).strip()
                        or str(cat_cfg.get("skip", "")).strip()
                    ):
                        continue
                    eff_keep, eff_skip = self._resolve_image_filter(
                        file_keep, file_skip, cat_cfg
                    )
                    print(
                        f"  图片设置 · 分类「{cat_title or '(无分类)'}」: "
                        + self._format_image_filter(eff_keep, eff_skip)
                    )

                # 读取 + 去重
                targets = read_targets_from_excel(fpath)
                json_url_groups, dup_count = group_targets_by_json_url(targets)
                total = len(json_url_groups)
                print(f"  共 {total} 个唯一商品（跳过 {dup_count} 个重复）")

                if total == 0:
                    print("  无商品，跳过")
                    self._emit_file_status(fi, "已跳过", "skipped")
                    current_results = None
                    current_output_file = None
                    continue

                # 一次性读取旧文件：读取失败时保护原文件，切换到 recovery。
                existing_rows = []
                if os.path.exists(output_file):
                    try:
                        _, existing_rows = load_xlsx(output_file)
                        print(f"  已读取旧文件: {len(existing_rows)} 条")
                    except Exception as exc:
                        self.recovery_count += 1
                        protected_file = output_file
                        output_file = build_recovery_xlsx_path(protected_file)
                        current_output_file = output_file
                        existing_rows = []
                        print(
                            f"  [保护] 旧 XLSX 读取失败，禁止覆盖原文件: "
                            f"{protected_file}"
                        )
                        print(f"  原因: {exc}")
                        print(f"  本轮数据将另存为: {output_file}")

                done_urls = set()
                for row in existing_rows:
                    href = row.get("link-href", "")
                    if href:
                        done_urls.add(str(href).strip())
                if done_urls:
                    print(f"  断点续传: 已完成 {len(done_urls)} 个")

                todo = {
                    url: group
                    for url, group in json_url_groups.items()
                    if group["original_url"] not in done_urls
                }
                if not todo:
                    print("  全部已完成")
                    all_results.extend(existing_rows)
                    self._emit_file_status(fi, "已完成", "success")
                    current_results = None
                    current_output_file = None
                    continue

                results = list(existing_rows)
                current_results = results

                self._emit_progress(
                    0,
                    len(todo),
                    f"[{fi + 1}/{file_count}] 正在启动浏览器 · {fname}",
                )

                chrome = None
                success_count = 0
                fail_count = 0
                file_exception = None

                try:
                    chrome = create_browser()
                    tab = chrome.latest_tab

                    for index, (json_url, group) in enumerate(todo.items(), 1):
                        if self.stop_event.is_set():
                            break

                        table_title = group["title"]
                        original_url = group["original_url"]

                        self._emit_progress(
                            index,
                            len(todo),
                            f"[{fi + 1}/{file_count}] {fname} · "
                            f"{table_title[:24]}",
                        )
                        # 每个商品用它所属分类实际生效的图片规则，并写进日志，
                        # 便于核对「设置是否按预期作用到了这个分类」。
                        item_keep, item_skip = self._resolve_image_filter(
                            file_keep,
                            file_skip,
                            category_filters.get(table_title),
                        )
                        print(
                            f"  [{index}/{len(todo)}] {table_title}"
                            f" · 图片 {self._format_image_filter(item_keep, item_skip)}"
                        )

                        retry_count = 0
                        item = None

                        while retry_count <= MAX_RETRIES:
                            if self.stop_event.is_set():
                                break
                            status, data = fetch_json(
                                tab, json_url, stop_event=self.stop_event
                            )
                            if status == 200 and data:
                                # 解析失败（含 styles1 超过 MAX_STYLES_LENGTH）
                                # 只跳过当前商品，不能让整个文件中断。
                                try:
                                    with product_parser.with_image_filter(
                                        item_keep, item_skip
                                    ):
                                        item = parse_product(
                                            data, table_title, original_url
                                        )
                                except product_parser.InvalidProductError as exc:
                                    print(f"    跳过当前商品: {exc}")
                                break

                            if status == 429:
                                retry_count += 1
                                if retry_count > MAX_RETRIES:
                                    print(
                                        f"    429 超过 {MAX_RETRIES} 次，放弃"
                                    )
                                    break
                                delay = RETRY_BASE_DELAY * retry_count
                                print(
                                    f"    429 第{retry_count}次，{delay}秒后重试..."
                                )
                                if self.stop_event.wait(delay):
                                    break
                            else:
                                retry_count += 1
                                if retry_count > MAX_RETRIES:
                                    print(
                                        f"    失败 {MAX_RETRIES} 次，放弃 "
                                        f"(status={status})"
                                    )
                                    break
                                delay = 5 * retry_count
                                print(
                                    f"    失败 (status={status})，{delay}秒后重试..."
                                )
                                if self.stop_event.wait(delay):
                                    break

                            if self.stop_event.wait(DELAY):
                                break

                        if self.stop_event.is_set():
                            break

                        if item:
                            results.append(item)
                            success_count += 1
                            print(f"    成功: {item['name'][:40]}")
                        else:
                            fail_count += 1

                        if success_count % SAVE_EVERY == 0 and success_count > 0:
                            save_xlsx(results, output_file)
                            print(f"    [自动保存] {len(results)} 条")

                        if self.stop_event.wait(DELAY):
                            break

                except Exception as exc:
                    file_exception = exc
                    print(f"  当前文件采集异常: {exc}")
                    logger.log_exc()

                finally:
                    # 浏览器关闭失败必须与数据保存彻底隔离。
                    if chrome is not None:
                        print("  正在关闭浏览器...")
                        try:
                            chrome.quit()
                        except Exception as exc:
                            print(f"  关闭浏览器失败（不影响保存）: {exc}")

                    # 无论正常完成、停止还是异常，都强制保存当前已采集结果。
                    if results:
                        try:
                            os.makedirs(
                                os.path.dirname(output_file), exist_ok=True
                            )
                            save_xlsx(results, output_file)
                            print(f"  文件保存: {output_file} ({len(results)} 条)")
                        except Exception as save_exc:
                            file_exception = save_exc
                            print(f"  [严重] 保存当前已采集结果失败: {save_exc}")
                            logger.log_exc()

                all_results.extend(results)
                grand_success += success_count
                grand_fail += fail_count
                self.failed_count = grand_fail
                current_results = None
                current_output_file = None

                if file_exception is not None:
                    raise file_exception

                if self.stop_event.is_set():
                    self._emit_file_status(fi, "已停止", "stopped")
                    break

                if file_exception is not None:
                    self._emit_file_status(fi, "发生异常", "failed")
                    # 保持“任务遇到未处理异常则结束”的总体语义，但数据已经先保存。
                    raise file_exception

                if fail_count > 0 and success_count == 0:
                    self._emit_file_status(fi, "采集失败", "failed")
                else:
                    self._emit_file_status(fi, "已完成", "success")

            if self.stop_event.is_set() and active_file_index is not None:
                for pending_index in range(
                    active_file_index + 1, len(self.task_files)
                ):
                    self._emit_file_status(pending_index, "未处理", "waiting")

            final_text = (
                f"总计：{grand_success} 成功，{grand_fail} 失败，"
                f"{len(all_results)} 条数据"
            )
            if self.stop_event.is_set():
                final_text = "任务已中断 · " + final_text
            self._emit_progress(
                grand_success + grand_fail,
                grand_success + grand_fail,
                final_text,
            )
            print(f"\n{final_text}")

            return all_results

        except Exception as exc:
            # 双保险：异常发生在当前文件保存 finally 之外时再尝试一次落盘。
            if current_results and current_output_file:
                try:
                    os.makedirs(
                        os.path.dirname(current_output_file), exist_ok=True
                    )
                    save_xlsx(current_results, current_output_file)
                    print(
                        f"[异常保护保存] {current_file_name}: "
                        f"{current_output_file} ({len(current_results)} 条)"
                    )
                except Exception as save_exc:
                    print(f"[严重] 异常保护保存失败: {save_exc}")

            print(f"发生异常: {exc}")
            logger.log_exc()
            if active_file_index is not None:
                self._emit_file_status(active_file_index, "发生异常", "failed")
            self._emit_progress(0, 0, f"任务异常：{exc}")
            raise