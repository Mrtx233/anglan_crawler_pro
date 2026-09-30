from __future__ import annotations

from urllib.parse import parse_qs, urlencode, urlparse, urlunparse
import json
import traceback
from ..config import LINK_MAX_RETRIES, RETRY_DELAY
from ..crawler.browser import create_browser
from ..crawler.links import collect_page_links
from ..pipeline.models import PipelineStage, TaskResult
from ..storage.paths import to_shopify_json_url
from ..storage.workbooks import load_existing_link_rows, write_link_workbooks


def collect_category_links(context, pages, max_pages, xpath, output_folder):
    all_results = load_existing_link_rows(output_folder)
    saved_links = {to_shopify_json_url(row["link"]) for row in all_results}
    domain_rows = {}
    for row in all_results:
        domain_rows.setdefault(urlparse(row["link"]).netloc, []).append(row)
    context.log(f"已读取已有 XLSX：{len(all_results)} 条唯一商品链接")
    chrome = create_browser(log=context.log)
    try:
        tab = chrome.latest_tab
        for category_index, (title, base_url) in enumerate(pages, 1):
            if context.stop_event.is_set():
                break
            context.emit(
                "progress",
                category_index - 1,
                len(pages),
                f"第一阶段 [{category_index}/{len(pages)}] {title}",
            )
            context.log(f"\n[{category_index}/{len(pages)}] {title}")
            seen_in_category = set()
            added_count = 0
            changed_domains = set()
            page_num = 1
            try:
                while not context.stop_event.is_set():
                    if max_pages and page_num > max_pages:
                        break
                    parsed = urlparse(base_url)
                    query = parse_qs(parsed.query, keep_blank_values=True)
                    query["page"] = [str(page_num)]
                    url = (
                        base_url
                        if page_num == 1
                        else urlunparse(
                            parsed._replace(query=urlencode(query, doseq=True), fragment="")
                        )
                    )
                    context.log(f"  第 {page_num} 页: {url}")
                    links = []
                    for retry in range(LINK_MAX_RETRIES + 1):
                        if context.stop_event.is_set():
                            break
                        if tab.get(url) is False:
                            raise TimeoutError(f"页面加载失败: {url}")
                        links = collect_page_links(tab, context.stop_event, xpath)
                        if links or page_num > 1 or retry == LINK_MAX_RETRIES:
                            break
                        context.log(f"  首页无链接，等待后重试 ({retry + 1}/{LINK_MAX_RETRIES})")
                        if context.stop_event.wait(RETRY_DELAY):
                            break
                    new_on_page = 0
                    for link in links:
                        key = to_shopify_json_url(link)
                        if key in seen_in_category:
                            continue
                        seen_in_category.add(key)
                        new_on_page += 1
                        if key not in saved_links:
                            saved_links.add(key)
                            row = {"title": title, "link": link}
                            all_results.append(row)
                            domain = urlparse(link).netloc
                            domain_rows.setdefault(domain, []).append(row)
                            changed_domains.add(domain)
                            added_count += 1
                    if not new_on_page:
                        context.log("  本页无本轮新增链接，结束该分类")
                        break
                    context.log(
                        f"  分类已识别 {len(seen_in_category)} 条，去除已有数据后新增 {added_count} 条"
                    )
                    page_num += 1
            finally:
                _save_links_index(output_folder, pages[:category_index])
                for domain in sorted(changed_domains):
                    context.checkpoints.pending[f"links:{domain}"] = (
                        write_link_workbooks,
                        (domain_rows[domain], output_folder),
                    )
                for domain in sorted(changed_domains):
                    key = f"links:{domain}"
                    writer, args = context.checkpoints.pending[key]
                    context.checkpoints.save(key, writer, *args)
                context.log(
                    f"  分类已保存：本分类新增 {added_count} 条，总计 {len(all_results)} 条"
                )
        return all_results
    finally:
        context.log("第一阶段：正在关闭浏览器...")
        try:
            chrome.quit()
        except Exception:
            context.log(traceback.format_exc())


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


def _save_links_index(output_folder, pages):
    """Persist domain -> category title -> URL mappings like the legacy collector."""
    from pathlib import Path

    path = Path(output_folder) / "links_index.json"
    index = {}
    if path.exists():
        try:
            index = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError, TypeError):
            index = {}
    for title, url in pages:
        parsed = urlparse(str(url).strip())
        if not parsed.netloc:
            continue
        bucket = index.setdefault(parsed.netloc, {})
        urls = bucket.setdefault(str(title).strip(), [])
        if isinstance(urls, str):
            urls = bucket[str(title).strip()] = [urls]
        if url not in urls:
            urls.append(url)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(_dump_compact_index(index), encoding="utf-8")


def run_link_task(context, config):
    rows = collect_category_links(
        context, config.pages, config.max_pages, config.xpath, config.folder
    )
    return TaskResult(PipelineStage.COLLECTING_LINKS, link_count=len(rows))
