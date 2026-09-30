# ======================== 采集核心层 ========================
#
# 职责: 分页采集业务逻辑。含分类文本解析、分页 URL 构造、
#       链接提取（HTML 正则 / XPath 两种模式）、采集主循环。
# 依赖: config、logger、browser_utils、file_utils。
# 禁止: 导入 tkinter —— 本模块必须能脱离 GUI 运行与单测。
#
# GUI 与本模块的通信方式:
#   1. CollectConfig（输入）
#   2. threading.Event（停止）
#   3. on_progress 回调（分类进度，可选）
#   4. on_save_report 回调（保存报告，可选）
#   5. on_count 回调（累计链接数，可选）
# ========================================================

import json
import re
import threading
import time
from typing import Callable, List, Optional
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse

import browser_utils
import config
import file_utils
import logger

ContextLostError = browser_utils.ContextLostError

ProgressCallback = Callable[[int, int, str], None]
# 保存报告回调签名: (SaveReport) -> None
SaveReportCallback = Callable[[object], None]
# 累计链接数回调签名: (已采集条数) -> None
CountCallback = Callable[[int], None]


# ======================== 输入解析 ========================
def parse_categories(text: str) -> List[config.CategoryItem]:
    """解析分类列表文本，返回 [CategoryItem(title, url), ...]。"""
    raw = (text or "").strip()
    if not raw:
        return []

    categories: List[config.CategoryItem] = []
    for line in raw.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue

        url_match = re.search(r"(https?://[^\s,\"\')\]]+)", line)
        if not url_match:
            continue
        url = url_match.group(1)

        title = ""
        url_start = line.find(url)
        if url_start > 0:
            title_part = line[:url_start].strip(" \t,(\"'")
            title = re.sub(r"[\",\)\]]+$", "", title_part).strip()

        if not title:
            title = urlparse(url).path.split("/")[-1]

        categories.append(config.CategoryItem(title=title, url=url))
    return categories


# ======================== 分页 URL ========================
def build_page_url(base_url: str, page_num: int) -> str:
    """安全设置或替换 page 参数，避免重复 ?page= 或 fragment 拼接错误。"""
    if page_num <= 1:
        return base_url

    parsed = urlparse(base_url)
    query = dict(parse_qsl(parsed.query, keep_blank_values=True))
    query["page"] = str(page_num)
    return urlunparse(parsed._replace(query=urlencode(query)))


# ======================== 链接提取：HTML 正则模式 ========================
def extract_product_links(html_text: str, base_url: str) -> List[str]:
    """从当前分页 HTML 提取 Shopify 商品链接并按出现顺序去重。"""
    base_parsed = urlparse(base_url)
    base = f"{base_parsed.scheme}://{base_parsed.netloc}"
    seen_handles = set()
    unique_handles: List[str] = []

    def add_handle(handle):
        handle = str(handle or "").strip().strip("/")
        if handle and handle not in seen_handles:
            seen_handles.add(handle)
            unique_handles.append(handle)

    href_pattern = r'href=["\'](?:https?://[^/"\']+)?(/products/[^"\'?#/]+)'
    for href in re.findall(href_pattern, html_text, re.IGNORECASE):
        add_handle(href.rstrip("/").split("/")[-1])

    meta_match = re.search(
        r"var\s+meta\s*=\s*(\{.+?\});\s*(?:\n|</script>)",
        html_text,
        re.DOTALL,
    )
    if meta_match:
        try:
            meta_data = json.loads(meta_match.group(1))
            for product in meta_data.get("products", []):
                add_handle(product.get("handle", ""))
        except (json.JSONDecodeError, TypeError, KeyError):
            pass

    for handle in re.findall(r'"handle"\s*:\s*"([^"]+)"', html_text):
        path_pattern = rf"/products/{re.escape(handle)}(?=[/?#\"'<>\s]|$)"
        if re.search(path_pattern, html_text, re.IGNORECASE):
            add_handle(handle)

    return [f"{base}/products/{handle}" for handle in unique_handles]


# ======================== 链接提取：XPath 模式 ========================
def extract_links_by_xpath(
    tab,
    xpath: str,
    stop_event: Optional[threading.Event] = None,
) -> List[str]:
    """用 XPath 提取商品详情页链接，只保留 /products/...。"""
    result = browser_utils.evaluate_xpath(tab, xpath, stop_event=stop_event)

    if isinstance(result, dict) and result.get("__xpath_error__"):
        raise ValueError(f"XPath 无效: {result['__xpath_error__']}")
    if not isinstance(result, list):
        return []

    base_parsed = urlparse(tab.url)
    base = f"{base_parsed.scheme}://{base_parsed.netloc}"
    seen = set()
    links: List[str] = []

    for raw in result:
        raw = str(raw).strip()
        if not raw:
            continue
        if raw.startswith("/"):
            full = f"{base}{raw}"
        elif raw.startswith("http"):
            full = raw
        else:
            continue

        parsed = urlparse(full)
        match = re.match(r"^/products/([^/?#]+)", parsed.path, re.IGNORECASE)
        if not match:
            continue
        normalized = f"{base}/products/{match.group(1)}"
        if normalized not in seen:
            seen.add(normalized)
            links.append(normalized)

    return links


# ======================== 单页采集 ========================
def collect_page_links(
    tab,
    stop_event: Optional[threading.Event] = None,
    xpath: str = "",
) -> List[str]:
    """采集当前分页并返回链接。

    停止策略统一由调用方（CollectorRunner）在返回后判定并放弃本页，
    本函数不再自行丢弃结果，避免同一契约在函数内与每个调用点各写一遍。
    """
    if stop_event is not None:
        if stop_event.wait(config.PAGE_SETTLE_WAIT):
            return []
    else:
        time.sleep(config.PAGE_SETTLE_WAIT)

    if xpath:
        return extract_links_by_xpath(tab, xpath, stop_event=stop_event)

    html_text = browser_utils.get_outer_html(tab, stop_event=stop_event)
    if not html_text:
        return []
    return extract_product_links(html_text, tab.url)


# ======================== 采集主循环 ========================
class CollectorRunner:
    """原 _run_task 的业务循环，与 tkinter 彻底解耦。

    用法:
        runner = CollectorRunner(cfg, stop_event, on_progress=fn,
                                 on_save_report=fn2, on_count=fn3)
        rows = runner.run()          # 阻塞执行，内部已完成落盘
        runner.interrupt()           # 其它线程调用，请求停止
        runner.last_report           # 最近一次 SaveReport
    """

    def __init__(
        self,
        cfg: config.CollectConfig,
        stop_event: Optional[threading.Event] = None,
        on_progress: Optional[ProgressCallback] = None,
        on_save_report: Optional[SaveReportCallback] = None,
        on_count: Optional[CountCallback] = None,
    ):
        self.cfg = cfg
        self.stop_event = stop_event or threading.Event()
        self.on_progress = on_progress
        self.on_save_report = on_save_report
        self.on_count = on_count
        self._tab = None
        # 最近一次保存报告；即使 run() 抛错也会在 finally 中赋值
        self.last_report = None

    # ---------- 对外控制 ----------
    def interrupt(self) -> None:
        """请求停止：置 event 并尽力中止当前页面加载。"""
        self.stop_event.set()
        browser_utils.interrupt_page(self._tab)

    # ---------- 主流程 ----------
    def run(self) -> List[dict]:
        """执行采集，返回 [{"title":..., "link":...}, ...]。

        无论正常完成、用户停止还是运行异常，都会在 finally 中调用
        file_utils.save_results() 落盘已提交的页结果，并触发
        on_save_report 回调（若提供）。
        """
        all_results: List[dict] = []
        seen_all_links = set()
        chrome = None

        try:
            logger.log(self.cfg.summary())
            logger.log("")

            chrome = browser_utils.create_browser(
                proxy=self.cfg.proxy,
                port_range=self.cfg.port_range,
                disable_images=self.cfg.disable_images,
            )
            tab = chrome.latest_tab
            self._tab = tab

            categories = self.cfg.categories
            for i, category in enumerate(categories, 1):
                if self.stop_event.is_set():
                    break
                self._collect_category(
                    index=i,
                    total=len(categories),
                    category=category,
                    tab=tab,
                    all_results=all_results,
                    seen_all_links=seen_all_links,
                )

        except Exception as e:
            logger.log(f"发生异常: {e}")
            logger.log_exc()

        finally:
            self._tab = None
            if chrome is not None:
                try:
                    chrome.quit()
                except Exception as e:
                    logger.log(f"关闭浏览器时出现异常（不影响保存）: {e}")

            # 无论正常完成、用户停止还是运行异常，都保存此前已完成分页的结果。
            report = file_utils.save_results(
                all_results, self.cfg.output_dir, self.cfg.categories
            )
            self.last_report = report

            if self.on_save_report is not None:
                try:
                    self.on_save_report(report)
                except Exception as e:
                    logger.log(f"保存报告回调异常（不影响结果）: {e}")

        return all_results

    # ---------- 单个分类 ----------
    def _collect_category(
        self,
        index: int,
        total: int,
        category: config.CategoryItem,
        tab,
        all_results: List[dict],
        seen_all_links: set,
    ) -> None:
        title = category.title
        base_url = category.url

        category_total = 0
        category_new_for_all = 0
        category_skipped = 0

        try:
            logger.log(f"[{index}/{total}] {title}")
            if self.on_progress is not None:
                self.on_progress(index, total, title)

            seen_in_category = set()
            page_num = 1

            while not self.stop_event.is_set():
                if self.cfg.max_pages and page_num > self.cfg.max_pages:
                    logger.log(f"  已达到最大页数 {self.cfg.max_pages}")
                    break

                url = build_page_url(base_url, page_num)
                logger.log(f"  第 {page_num} 页: {url}")

                links = self._fetch_page_with_recovery(tab, url)
                if links is None:
                    break

                if self.stop_event.is_set():
                    logger.log("  当前页已放弃")
                    break

                new_links = [
                    link for link in links if link not in seen_in_category
                ]

                if not new_links:
                    if page_num == 1:
                        new_links = self._retry_first_page(
                            tab, url, seen_in_category
                        )
                        if self.stop_event.is_set():
                            logger.log("  当前页已放弃")
                            break
                        if not new_links:
                            logger.log("  跳过该分类")
                            break
                    else:
                        logger.log("  无新链接，该分类采集完毕")
                        break

                # 这里不再重复检查 stop_event：317 行与 330 行的检查已覆盖
                # 所有会阻塞的取页路径，此后到提交之间没有 I/O，无需第三次判定。
                page_new_for_all = 0
                page_skipped = 0
                for link in new_links:
                    seen_in_category.add(link)
                    category_total += 1

                    if link in seen_all_links:
                        category_skipped += 1
                        page_skipped += 1
                        continue
                    seen_all_links.add(link)
                    all_results.append({"title": title, "link": link})
                    category_new_for_all += 1
                    page_new_for_all += 1

                logger.log(
                    f"  新增 {len(new_links)} 条（分类累计 {category_total} 条，"
                    f"全局新增 {page_new_for_all} 条，跨分类跳过 {page_skipped} 条）"
                )
                if self.on_count is not None:
                    try:
                        self.on_count(len(all_results))
                    except Exception as e:
                        logger.log(f"  链接计数回调异常（不影响采集）: {e}")
                page_num += 1

            logger.log(
                f"  共采集 {category_total} 条（新增 {category_new_for_all} 条，"
                f"跨分类去重跳过 {category_skipped} 条）"
            )
            logger.log("")

        except Exception as e:
            logger.log(f"  分类 [{title}] 出错: {e}")
            # 与 run() 的外层 except 保持一致：保留堆栈，否则编程错误
            # 和普通的页面故障在日志里无法区分。
            logger.log_exc()
            if category_total:
                logger.log(
                    f"  已保留该分类此前完成分页的 {category_total} 条链接"
                )
            logger.log("")

    # ---------- 单页加载（含上下文丢失重试） ----------
    def _fetch_page_with_recovery(self, tab, url: str) -> Optional[List[str]]:
        """加载并采集一页；返回 None 表示页面上下文彻底丢失、应结束该分类。"""
        try:
            tab.get(url)
            if self.stop_event.is_set():
                logger.log("  当前页已放弃")
                return []
            return collect_page_links(
                tab, stop_event=self.stop_event, xpath=self.cfg.xpath
            )
        except ContextLostError:
            if self.stop_event.is_set():
                logger.log("  当前页已放弃")
                return []
            logger.log("  页面上下文丢失，快速重试一次...")
            if self.stop_event.wait(config.CONTEXT_LOST_RETRY_DELAY):
                return []
            try:
                tab.get(url)
                if self.stop_event.is_set():
                    logger.log("  当前页已放弃")
                    return []
                return collect_page_links(
                    tab, stop_event=self.stop_event, xpath=self.cfg.xpath
                )
            except ContextLostError:
                logger.log("  页面未恢复，结束该分类")
                return None

    # ---------- 首页空结果重试 ----------
    def _retry_first_page(self, tab, url: str, seen_in_category: set) -> List[str]:
        """首页无链接时重试 MAX_RETRIES 次，返回首次拿到的新链接（可能为空）。"""
        for retry in range(1, config.MAX_RETRIES + 1):
            if self.stop_event.is_set():
                break
            logger.log(f"  首页无链接，第 {retry}/{config.MAX_RETRIES} 次重试...")
            if self.stop_event.wait(config.RETRY_DELAY):
                break

            links = self._fetch_page_with_recovery(tab, url)
            # None 表示上下文丢失且未恢复（_fetch_page_with_recovery 已记日志），
            # 与主路径一致立即结束该分类；不能折叠成 []，否则会把「页面已死」
            # 误判为「本页无链接」而白跑满重试次数。
            if links is None:
                break
            if self.stop_event.is_set():
                break

            new_links = [link for link in links if link not in seen_in_category]
            if new_links:
                logger.log(f"  重试成功: {len(new_links)} 条")
                return new_links

        return []