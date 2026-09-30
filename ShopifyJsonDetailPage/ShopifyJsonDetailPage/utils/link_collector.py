"""
DrissionPage 采集 collection 页面商品链接。

spider 配置 PAGES + XPATH_LINK 后，在 start_requests 阶段调用本模块，
先通过浏览器自动化收集商品链接，再将结果作为 targets 喂给 Scrapy。
"""

import os
import random
import time
from urllib.parse import urljoin, urlparse

import openpyxl
from DrissionPage import Chromium, ChromiumOptions
from lxml import html as lxml_html


# ── 浏览器 ──────────────────────────────────────────────────

def _create_browser(proxy="http://127.0.0.1:7897"):
    co = ChromiumOptions()
    co.incognito(on_off=True)
    co.set_argument(f"--proxy-server={proxy}")
    port = random.randint(9222, 9322)
    co.set_local_port(port)

    for attempt in range(3):
        try:
            chrome = Chromium(co)
            print(f"[browser] 启动成功，端口 {port}")
            return chrome
        except Exception as exc:
            print(f"[browser] 端口 {port} 失败 (尝试 {attempt + 1}/3): {exc}")
            port = random.randint(9222, 9322)
            co.set_local_port(port)
    raise RuntimeError("浏览器启动失败，已重试 3 次")


# ── 页面操作 ─────────────────────────────────────────────────

def _scroll_page(tab, pause=1.5, max_scrolls=30):
    """滚动到底部触发懒加载，高度不再变化时停止。"""
    from DrissionPage.errors import ContextLostError

    # 等待页面稳定（防止 redirect 导致 context lost）
    time.sleep(2)
    for wait_attempt in range(3):
        try:
            last = tab.run_js("return document.body.scrollHeight")
            break
        except ContextLostError:
            if wait_attempt < 2:
                print(f"  [scroll] 页面刷新中，等待重试 ({wait_attempt + 1}/3)...")
                time.sleep(3)
                tab.refresh()
                time.sleep(2)
            else:
                print("  [scroll] 页面无法稳定，跳过滚动")
                return

    for _ in range(max_scrolls):
        try:
            tab.run_js("window.scrollTo(0, document.body.scrollHeight)")
            time.sleep(pause)
            new = tab.run_js("return document.body.scrollHeight")
        except ContextLostError:
            print("  [scroll] 页面刷新，停止滚动")
            break
        if new == last:
            break
        last = new


def _get_tree(tab):
    from DrissionPage.errors import ContextLostError
    for attempt in range(3):
        try:
            raw = tab.run_js("return document.documentElement.outerHTML;")
            return lxml_html.fromstring(raw)
        except ContextLostError:
            if attempt < 2:
                print(f"  [tree] 页面刷新中，等待重试 ({attempt + 1}/3)...")
                time.sleep(3)
                tab.refresh()
                time.sleep(2)
            else:
                raise


def _resolve_url(href, page_url):
    if not href:
        return ""
    if href.startswith("//"):
        return "https:" + href
    if href.startswith("http"):
        return href
    return urljoin(page_url, href)


def _extract_links(tree, xpath, page_url):
    els = tree.xpath(xpath) if xpath else []
    links, seen = [], set()
    for el in els:
        href = el.strip() if isinstance(el, str) else ""
        if not href and hasattr(el, "get"):
            href = el.get("href", "").strip()
        href = _resolve_url(href, page_url)
        if href and href not in seen:
            seen.add(href)
            links.append(href)
    return links


# ── 保存详细链接 XLSX ────────────────────────────────────────

def _save_links_xlsx(results, spider_name, output_dir):
    """
    保存采集到的链接到 {spider}_详细链接.xlsx。
    results: list[dict] with keys "title", "link"
    """
    os.makedirs(output_dir, exist_ok=True)
    path = os.path.join(output_dir, f"{spider_name}_详细链接.xlsx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = spider_name[:31]
    ws.append(["title", "link"])
    for r in results:
        ws.append([r.get("title", ""), r.get("link", "")])
    wb.save(path)
    print(f"[links] 已保存: {path} ({len(results)} 条)")
    return path


# ── 主入口 ──────────────────────────────────────────────────

def collect_links(pages, xpath_link, spider_name, output_dir,
                  max_pages=5, start_page=1, max_retries=3,
                  retry_delay=5, page_load_wait=2):
    """
    从 PAGES 列表逐页采集商品链接，返回 (title, url) 元组列表。

    参数:
        pages:        [(title, collection_url), ...]
        xpath_link:   商品链接 XPath
        spider_name:  spider 名称（用于输出文件名）
        output_dir:   详细链接 XLSX 保存目录
        max_pages:    每个分类最多采集页数
        start_page:   起始页码
        max_retries:  首页无链接时重试次数
        retry_delay:  重试间隔秒数
        page_load_wait: 页面加载后额外等待秒数
    """
    if not pages or not xpath_link:
        print("[links] PAGES 或 XPATH_LINK 未配置")
        return []

    print(f"[links] 共 {len(pages)} 个页面, XPath: {xpath_link}")
    chrome = _create_browser()
    tab = chrome.latest_tab
    all_results = []

    try:
        for i, (title, base_url) in enumerate(pages, 1):
            print(f"[links] [{i}/{len(pages)}] {title}")
            page_links = []
            seen_in_cat = set()
            page_num = start_page

            while page_num <= max_pages:
                # 拼接分页 URL
                if page_num == 1:
                    url = base_url
                else:
                    sep = "&" if "?" in base_url else "?"
                    url = f"{base_url}{sep}page={page_num}"

                print(f"  第 {page_num} 页: {url}")
                tab.get(url)
                time.sleep(page_load_wait)
                _scroll_page(tab)

                tree = _get_tree(tab)
                links = _extract_links(tree, xpath_link, tab.url)
                new_links = [l for l in links if l not in seen_in_cat]

                if not new_links:
                    if page_num == start_page:
                        # 首页无链接 → 重试
                        retried = False
                        for attempt in range(1, max_retries + 1):
                            print(f"  首页无链接，重试 {attempt}/{max_retries} ({retry_delay}s)...")
                            time.sleep(retry_delay)
                            tab.get(url)
                            time.sleep(page_load_wait)
                            _scroll_page(tab)
                            tree = _get_tree(tab)
                            links = _extract_links(tree, xpath_link, tab.url)
                            new_links = [l for l in links if l not in seen_in_cat]
                            if new_links:
                                print(f"  重试成功，{len(new_links)} 条新链接")
                                retried = True
                                break
                        if not retried:
                            print(f"  重试 {max_retries} 次仍无链接，跳过该分类")
                            break
                    else:
                        print(f"  第 {page_num} 页无新链接，分类完毕")
                        break

                for link in new_links:
                    seen_in_cat.add(link)
                    page_links.append(link)
                print(f"  新增 {len(new_links)} 条（累计 {len(page_links)} 条）")
                page_num += 1

            for link in page_links:
                all_results.append({"title": title, "link": link})
            print(f"  共 {len(page_links)} 条\n")

    finally:
        chrome.quit()

    if all_results:
        _save_links_xlsx(all_results, spider_name, output_dir)
    else:
        print("[links] 未采集到任何链接")

    # 返回 (title, url) 元组列表，供 spider 使用
    return [(r["title"], r["link"]) for r in all_results]
