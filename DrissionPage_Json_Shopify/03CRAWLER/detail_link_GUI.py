# ======================== 详情页链接采集 GUI ========================
#
# 用途: 从 Shopify 列表页提取产品链接
# 原理: DrissionPage 浏览器加载页面，从 <script> JSON 中提取 handle
#
# 界面:
#   - 分类列表文本框（每行: 标题, URL）
#   - 最大页数输入框（留空=不限制）
#   - 保存路径选择
#   - 开始/停止按钮 + 实时日志
#
# 输出: XLSX（title, link）
# ========================================================================

import json
import os
import random
import re
import sys
import threading
import time
from urllib.parse import urlparse

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import openpyxl
from DrissionPage import Chromium, ChromiumOptions
from DrissionPage.errors import ContextLostError

# ======================== 样式配置 ========================
UI_COLORS = {
    "background": "#F4F7FB",
    "card": "#FFFFFF",
    "card_alt": "#F8FAFC",
    "border": "#E2E8F0",
    "text": "#0F172A",
    "text_secondary": "#475569",
    "text_muted": "#94A3B8",
    "primary": "#2563EB",
    "primary_hover": "#1D4ED8",
    "success": "#16A34A",
    "danger": "#DC2626",
    "danger_light": "#FEF2F2",
    "log_background": "#0F172A",
    "log_panel": "#111827",
    "log_text": "#D1D5DB",
}

UI_FONT = "Microsoft YaHei UI"
MONO_FONT = "Consolas"

MAX_RETRIES = 3
RETRY_DELAY = 5
PAGE_LOAD_WAIT = 1.5

# 无限滚动配置
INFINITE_SCROLL_WAIT = 1.2
INFINITE_SCROLL_STABLE_ROUNDS = 5
INFINITE_SCROLL_MAX_ROUNDS = 80


# ======================== 滚动加载 ========================
def scroll_to_bottom(tab):
    """滚动到底部，等待5秒，连续3次高度不变则停止。"""
    try:
        last_height = tab.run_js("return document.body.scrollHeight;")
    except ContextLostError:
        time.sleep(3)
        last_height = tab.run_js("return document.body.scrollHeight;")

    no_change = 0
    while no_change < 3:
        try:
            tab.run_js("window.scrollTo(0, document.body.scrollHeight);")
        except ContextLostError:
            print("  页面刷新，等待恢复...")
            time.sleep(3)
            try:
                tab.run_js("window.scrollTo(0, document.body.scrollHeight);")
            except ContextLostError:
                print("  页面未恢复，跳过滚动")
                return
        time.sleep(5)

        try:
            new_height = tab.run_js("return document.body.scrollHeight;")
        except ContextLostError:
            time.sleep(3)
            try:
                new_height = tab.run_js("return document.body.scrollHeight;")
            except ContextLostError:
                return

        if new_height > last_height:
            last_height = new_height
            no_change = 0
        else:
            no_change += 1


# ======================== 浏览器初始化 ========================
def create_browser():
    co = ChromiumOptions()
    co.incognito(on_off=True)
    co.set_argument("--proxy-server=http://127.0.0.1:7897")
    random_port = random.randint(9222, 9322)
    co.set_local_port(random_port)

    for attempt in range(3):
        try:
            chrome = Chromium(co)
            print(f"浏览器启动成功，端口: {random_port}")
            return chrome
        except Exception as e:
            print(f"端口 {random_port} 启动失败 (尝试 {attempt + 1}/3): {e}")
            random_port = random.randint(9222, 9322)
            co.set_local_port(random_port)
    raise RuntimeError("无法启动浏览器，已重试多次")


# ======================== 提取产品链接 ========================
def extract_product_links(html_text, base_url):
    """从页面提取产品链接，合并三种来源确保完整：
    meta JSON + handle正则 + DOM链接，去重后返回。
    """
    base_parsed = urlparse(base_url)
    base = f"{base_parsed.scheme}://{base_parsed.netloc}"
    seen_handles = set()
    unique_handles = []

    # 方式1: 解析 var meta.products
    meta_match = re.search(
        r"var\s+meta\s*=\s*(\{.+?\});\s*(?:\n|</script>)",
        html_text,
        re.DOTALL,
    )
    if meta_match:
        try:
            meta_data = json.loads(meta_match.group(1))
            for product in meta_data.get("products", []):
                handle = str(product.get("handle", "")).strip()
                if handle and handle not in seen_handles:
                    seen_handles.add(handle)
                    unique_handles.append(handle)
        except (json.JSONDecodeError, TypeError, KeyError):
            pass

    # 方式2: 正则提取所有 handle（捕获懒加载追加的 JSON 数据）
    for h in re.findall(r'"handle"\s*:\s*"([^"]+)"', html_text):
        if h not in seen_handles:
            seen_handles.add(h)
            unique_handles.append(h)

    # 方式3: 从 DOM 链接提取 /products/xxx（捕获 JS 渲染的产品卡片）
    href_pattern = r'href=["\'](?:https?://[^/"\']+)?(/products/[^"\'?#]+)'
    for href in re.findall(href_pattern, html_text, re.IGNORECASE):
        h = href.rstrip("/").split("/")[-1].strip()
        if h and h not in seen_handles:
            seen_handles.add(h)
            unique_handles.append(h)

    return [f"{base}/products/{h}" for h in unique_handles]


def _run_js_with_retry(tab, script, default=None, retries=2):
    """执行 JavaScript，并处理页面上下文临时丢失。"""
    for attempt in range(retries + 1):
        try:
            result = tab.run_js(script)
            return default if result is None else result
        except ContextLostError:
            if attempt >= retries:
                return default
            time.sleep(2)
        except Exception:
            if attempt >= retries:
                return default
            time.sleep(1)
    return default


def is_infinite_scroll_page(tab):
    """判断当前分类页是否使用无限滚动商品列表。"""
    result = _run_js_with_retry(
        tab,
        r"""
        return (() => {
            const lists = Array.from(document.querySelectorAll('product-list'));
            if (lists.some(el => {
                const value = String(el.getAttribute('enable-infinite-scroll') || '').toLowerCase();
                return value === 'true'
                    || (value === '' && el.hasAttribute('enable-infinite-scroll'));
            })) {
                return true;
            }
            return Boolean(
                document.querySelector('.ProductList--infinite, [data-infinite-scroll="true"]')
            );
        })();
        """,
        default=False,
    )
    return bool(result)


def extract_rendered_product_links(tab):
    """从商品列表中提取已渲染链接，包括开放的 Shadow DOM。"""
    links = _run_js_with_retry(
        tab,
        r"""
        return (() => {
            const found = new Set();
            const visited = new Set();

            function addProductLink(href) {
                if (!href) return;
                try {
                    const url = new URL(href, location.origin);
                    const match = url.pathname.match(/^\/products\/([^\/?#]+)/i);
                    if (!match) return;
                    found.add(`${url.origin}/products/${match[1]}`);
                } catch (error) {}
            }

            function scanRoot(root) {
                if (!root || visited.has(root)) return;
                visited.add(root);

                if (root.matches && root.matches('a[href*="/products/"]')) {
                    addProductLink(root.getAttribute('href') || root.href);
                }
                if (!root.querySelectorAll) return;

                root.querySelectorAll('a[href*="/products/"]').forEach(anchor => {
                    addProductLink(anchor.getAttribute('href') || anchor.href);
                });

                root.querySelectorAll('*').forEach(element => {
                    if (element.shadowRoot) scanRoot(element.shadowRoot);
                });
            }

            const selectors = [
                'product-list',
                '#ProductList',
                '.ProductList--grid',
                '[data-product-list]',
                '[data-product-grid]',
                '#product-grid',
                '.collection-product-grid'
            ];
            const roots = [];
            selectors.forEach(selector => {
                document.querySelectorAll(selector).forEach(element => roots.push(element));
            });

            if (roots.length) roots.forEach(scanRoot);
            else scanRoot(document);

            return Array.from(found);
        })();
        """,
        default=[],
    )

    if not isinstance(links, list):
        return []

    base_parsed = urlparse(tab.url)
    base = f"{base_parsed.scheme}://{base_parsed.netloc}"
    normalized = []
    seen = set()

    for link in links:
        parsed = urlparse(str(link))
        match = re.match(r"^/products/([^/?#]+)", parsed.path, re.IGNORECASE)
        if not match:
            continue
        normalized_link = f"{base}/products/{match.group(1)}"
        if normalized_link not in seen:
            seen.add(normalized_link)
            normalized.append(normalized_link)

    return normalized


def _click_load_more_in_product_list(tab):
    """尝试点击商品列表内部的 Load More / Show More 按钮。"""
    clicked_text = _run_js_with_retry(
        tab,
        r"""
        return (() => {
            const visited = new Set();
            const accepted = [
                'load more',
                'load more products',
                'show more',
                'show more products',
                'view more',
                'view more products'
            ];

            function isVisible(element) {
                const rect = element.getBoundingClientRect();
                const style = getComputedStyle(element);
                return rect.width > 0 && rect.height > 0
                    && style.display !== 'none'
                    && style.visibility !== 'hidden'
                    && !element.disabled;
            }

            function search(root) {
                if (!root || visited.has(root) || !root.querySelectorAll) return '';
                visited.add(root);

                for (const control of root.querySelectorAll('button, a, [role="button"]')) {
                    const text = String(control.textContent || '').trim().toLowerCase();
                    if (accepted.includes(text) && isVisible(control)) {
                        control.click();
                        return text;
                    }
                }

                for (const element of root.querySelectorAll('*')) {
                    if (element.shadowRoot) {
                        const result = search(element.shadowRoot);
                        if (result) return result;
                    }
                }
                return '';
            }

            const list = document.querySelector(
                'product-list, #ProductList, .ProductList--infinite'
            );
            return search(list || document);
        })();
        """,
        default="",
    )
    return str(clicked_text or "")


def _scroll_product_list_to_end(tab):
    """把视口推进到商品列表末端，同时处理内部可滚动容器。"""
    return _run_js_with_retry(
        tab,
        r"""
        return (() => {
            const visited = new Set();

            function scrollNested(root) {
                if (!root || visited.has(root) || !root.querySelectorAll) return;
                visited.add(root);

                root.querySelectorAll('*').forEach(element => {
                    if (element.shadowRoot) scrollNested(element.shadowRoot);
                    try {
                        const style = getComputedStyle(element);
                        const overflowY = style.overflowY;
                        if (
                            element.scrollHeight > element.clientHeight + 20
                            && (overflowY === 'auto' || overflowY === 'scroll')
                        ) {
                            element.scrollTop = element.scrollHeight;
                        }
                    } catch (error) {}
                });
            }

            const list = document.querySelector(
                'product-list, #ProductList, .ProductList--infinite, [data-product-list]'
            );

            if (list) {
                if (list.shadowRoot) scrollNested(list.shadowRoot);
                scrollNested(list);

                const listBottom = window.scrollY + list.getBoundingClientRect().bottom;
                const maxScroll = Math.max(
                    0,
                    document.documentElement.scrollHeight - window.innerHeight
                );
                const target = Math.min(
                    maxScroll,
                    Math.max(
                        window.scrollY + window.innerHeight * 0.8,
                        listBottom - window.innerHeight * 0.7
                    )
                );
                window.scrollTo(0, target);
            } else {
                const maxScroll = Math.max(
                    0,
                    document.documentElement.scrollHeight - window.innerHeight
                );
                window.scrollTo(
                    0,
                    Math.min(maxScroll, window.scrollY + window.innerHeight * 0.8)
                );
            }
            return true;
        })();
        """,
        default=False,
    )


def _get_scroll_metrics(tab):
    metrics = _run_js_with_retry(
        tab,
        r"""
        return (() => {
            const list = document.querySelector(
                'product-list, #ProductList, .ProductList--infinite, [data-product-list]'
            );
            const targetBottom = list
                ? window.scrollY + list.getBoundingClientRect().bottom
                : document.documentElement.scrollHeight;
            return {
                documentHeight: document.documentElement.scrollHeight,
                targetBottom: targetBottom,
                atTargetEnd: window.scrollY + window.innerHeight >= targetBottom - 100
            };
        })();
        """,
        default={},
    )
    return metrics if isinstance(metrics, dict) else {}


def collect_infinite_scroll_links(tab, stop_event=None):
    """持续触发无限滚动，直到商品数量与列表高度稳定。"""
    seen = set()
    ordered_links = []
    stable_rounds = 0
    last_document_height = 0
    last_target_bottom = 0

    _run_js_with_retry(tab, "window.scrollTo(0, 0); return true;", default=False)
    time.sleep(0.5)

    for _round_no in range(1, INFINITE_SCROLL_MAX_ROUNDS + 1):
        if stop_event is not None and stop_event.is_set():
            break

        before_count = len(ordered_links)
        before_metrics = _get_scroll_metrics(tab)

        for link in extract_rendered_product_links(tab):
            if link not in seen:
                seen.add(link)
                ordered_links.append(link)

        if len(ordered_links) > before_count:
            print(
                f"  无限滚动加载: +{len(ordered_links) - before_count} 条"
                f"（累计 {len(ordered_links)} 条）"
            )

        clicked_text = _click_load_more_in_product_list(tab)
        if clicked_text:
            print(f"  点击加载按钮: {clicked_text}")

        _scroll_product_list_to_end(tab)
        time.sleep(INFINITE_SCROLL_WAIT)

        after_scan_start = len(ordered_links)
        for link in extract_rendered_product_links(tab):
            if link not in seen:
                seen.add(link)
                ordered_links.append(link)

        if len(ordered_links) > after_scan_start:
            print(
                f"  无限滚动加载: +{len(ordered_links) - after_scan_start} 条"
                f"（累计 {len(ordered_links)} 条）"
            )

        after_metrics = _get_scroll_metrics(tab)
        document_height = int(after_metrics.get("documentHeight") or 0)
        target_bottom = int(after_metrics.get("targetBottom") or 0)
        at_target_end = bool(after_metrics.get("atTargetEnd"))

        count_grew = len(ordered_links) > before_count
        height_grew = (
            document_height
            > max(last_document_height, int(before_metrics.get("documentHeight") or 0))
            or target_bottom
            > max(last_target_bottom, int(before_metrics.get("targetBottom") or 0))
        )

        if at_target_end and not count_grew and not height_grew and not clicked_text:
            stable_rounds += 1
        else:
            stable_rounds = 0

        last_document_height = max(last_document_height, document_height)
        last_target_bottom = max(last_target_bottom, target_bottom)

        if stable_rounds >= INFINITE_SCROLL_STABLE_ROUNDS:
            print(f"  无限滚动已稳定，共识别 {len(ordered_links)} 条商品链接")
            break
    else:
        print(
            f"  已达到最大滚动轮数 {INFINITE_SCROLL_MAX_ROUNDS}，"
            f"当前识别 {len(ordered_links)} 条"
        )

    _run_js_with_retry(tab, "window.scrollTo(0, 0); return true;", default=False)
    return ordered_links


def extract_links_by_xpath(tab, xpath):
    """用 XPath 从页面提取产品链接，返回规范化后的链接列表。"""
    result = _run_js_with_retry(
        tab,
        f"""
        return (() => {{
            const xpath = {json.dumps(xpath)};
            const results = [];
            try {{
                const snapshot = document.evaluate(
                    xpath, document, null,
                    XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null
                );
                for (let i = 0; i < snapshot.snapshotLength; i++) {{
                    const node = snapshot.snapshotItem(i);
                    const val = node.nodeType === 2 ? node.value : (node.href || node.textContent || '');
                    if (val) results.push(val.trim());
                }}
            }} catch(e) {{
                console.error('XPath error:', e);
            }}
            return results;
        }})();
        """,
        default=[],
    )
    if not isinstance(result, list):
        return []

    base_parsed = urlparse(tab.url)
    base = f"{base_parsed.scheme}://{base_parsed.netloc}"
    seen = set()
    links = []
    for raw in result:
        raw = str(raw).strip()
        if not raw:
            continue
        # 相对路径或绝对路径 → 完整 URL
        if raw.startswith("/"):
            full = f"{base}{raw}"
        elif raw.startswith("http"):
            full = raw
        else:
            continue
        parsed = urlparse(full)
        match = re.match(r"^/products/([^/?#]+)", parsed.path, re.IGNORECASE)
        if match:
            norm = f"{base}/products/{match.group(1)}"
        else:
            norm = f"{parsed.scheme}://{parsed.netloc}{parsed.path.rstrip('/')}"
        if norm not in seen:
            seen.add(norm)
            links.append(norm)
    return links


def collect_page_links(tab, stop_event=None, xpath=""):
    """采集当前页面链接，返回 (links, 是否为无限滚动页面)。
    XPath 模式：用 XPath 提取。
    非 XPath 模式：滚动后用 handle 提取（与 handle.py 逻辑一致）。
    """
    time.sleep(PAGE_LOAD_WAIT)

    if xpath:
        # XPath 模式：滚动加载后用 XPath 提取，不回退
        scroll_to_bottom(tab)
        links = extract_links_by_xpath(tab, xpath)
        return links, False

    # Handle 模式：滚动加载后提取 handle
    scroll_to_bottom(tab)

    html_text = _run_js_with_retry(
        tab,
        "return document.documentElement.outerHTML;",
        default="",
    )
    if not html_text:
        return [], False

    base_parsed = urlparse(tab.url)
    base = f"{base_parsed.scheme}://{base_parsed.netloc}"

    # 提取所有 handle
    all_handles = re.findall(r'"handle"\s*:\s*"([^"]+)"', html_text)

    # 过滤：只保留 HTML 中存在 /products/{handle} 路径的
    seen = set()
    links = []
    for h in all_handles:
        if h in seen:
            continue
        if f"/products/{h}" in html_text:
            seen.add(h)
            links.append(f"{base}/products/{h}")

    return links, False


# ======================== 文本重定向 ========================
class TextRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, string):
        if not string:
            return
        try:
            self.text_widget.after(0, self._write, string)
        except (tk.TclError, RuntimeError):
            pass

    def _write(self, string):
        try:
            self.text_widget.config(state=tk.NORMAL)
            self.text_widget.insert(tk.END, string)
            self.text_widget.see(tk.END)
            self.text_widget.config(state=tk.DISABLED)
        except tk.TclError:
            pass

    def flush(self):
        pass


# ======================== GUI ========================
def load_existing_domain_rows(file_path):
    """读取已存在的 {domain}.xlsx，返回按文件顺序去重的 {link: title} 映射。

    文件不存在或读取失败时返回空字典（失败即重新生成，不阻断采集流程）。
    """
    existing = {}
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
        print(f"读取已有文件失败，将重新生成: {e}")
        return {}
    return existing


LINK_INDEX_FILENAME = "links_index.json"


def update_link_index(output_path, pages):
    """把分类列表的 域名→分类标题→分类URL 映射记录进 {output_path}/links_index.json。

    结构: {域名: {分类标题: [分类URL, ...]}}
    已存在的文件会被读入并合并，同一分类下的 URL 按出现顺序追加去重，实现可追加。
    分类标题用列表而非单值：同域名下允许重名分类（如几个都叫 Classic 的列表页），
    URL 全部保留，不会因标题相同而丢失。
    """
    json_path = os.path.join(output_path, LINK_INDEX_FILENAME)
    index = {}
    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            index = _normalize_link_index(loaded)
        except (ValueError, TypeError, OSError):
            print(f"读取已有链接记录失败，将重建: {json_path}")

    added = 0
    title_counts = {}
    for title, url in pages:
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
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=2)
    print(f"已更新链接记录: {json_path}（新增 {added} 个分类）")

    # 重名分类的 xlsx title 列无法区分来源列表页，索引里已按顺序保留全部 URL，
    # 这里额外提示一次，方便操作员判断是否需要改成不同的分类名。
    repeated = [
        f"{domain} / {title} × {count}"
        for (domain, title), count in title_counts.items()
        if count > 1
    ]
    if repeated:
        print("提示: 以下分类标题在同一域名下重复，索引已保留全部 URL: " + "；".join(repeated))
    return json_path


def _normalize_link_index(loaded):
    """兼容旧版 {域名: {分类标题: 分类URL}}：标量一律升级为单元素列表。"""
    index = {}
    if not isinstance(loaded, dict):
        return index
    for domain, titles in loaded.items():
        if not isinstance(titles, dict):
            continue
        bucket = {}
        for title, value in titles.items():
            if isinstance(value, dict) or value is None:
                continue
            urls = []
            for raw in (value if isinstance(value, list) else [value]):
                raw = str(raw or "").strip()
                if raw:
                    urls.append(raw)
            if urls:
                bucket[str(title)] = urls
        if bucket:
            index[str(domain)] = bucket
    return index


class LinkCollectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("详情页链接采集")
        self.root.geometry("960x780")
        self.root.minsize(860, 680)
        self.root.configure(bg=UI_COLORS["background"])

        self.is_running = False
        self.stop_event = threading.Event()

        self.max_pages_var = tk.StringVar(value="")  # 留空=不限制
        self.xpath_var = tk.StringVar()
        self.output_var = tk.StringVar()

        self._build_ui()
        sys.stdout = TextRedirector(self.log_text)
        sys.stderr = sys.stdout

    def _build_ui(self):
        main = tk.Frame(self.root, bg=UI_COLORS["background"])
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=16)

        # ── 标题 ──
        header = tk.Frame(main, bg=UI_COLORS["background"])
        header.pack(fill=tk.X, pady=(0, 12))
        tk.Label(
            header, text="详情页链接采集",
            bg=UI_COLORS["background"], fg=UI_COLORS["text"],
            font=(UI_FONT, 18, "bold"),
        ).pack(side=tk.LEFT)
        tk.Label(
            header, text="从 Shopify 列表页自动提取产品链接",
            bg=UI_COLORS["background"], fg=UI_COLORS["text_secondary"],
            font=(UI_FONT, 9),
        ).pack(side=tk.LEFT, padx=(16, 0), pady=(8, 0))

        # ── 控制区 ──
        ctrl_card = tk.Frame(main, bg=UI_COLORS["card"],
                             highlightbackground=UI_COLORS["border"],
                             highlightthickness=1, bd=0)
        ctrl_card.pack(fill=tk.X, pady=(0, 10))
        ctrl_inner = tk.Frame(ctrl_card, bg=UI_COLORS["card"])
        ctrl_inner.pack(fill=tk.X, padx=16, pady=12)

        _entry_style = dict(
            bg=UI_COLORS["card_alt"], fg=UI_COLORS["text"],
            insertbackground=UI_COLORS["primary"],
            font=(UI_FONT, 10), relief=tk.FLAT, bd=0,
            highlightthickness=1, highlightbackground=UI_COLORS["border"],
            highlightcolor=UI_COLORS["primary"],
        )

        # 第一行: 保存路径 + XPath + 最大页数
        row1 = tk.Frame(ctrl_inner, bg=UI_COLORS["card"])
        row1.pack(fill=tk.X, pady=(0, 8))

        # 保存路径
        col_path = tk.Frame(row1, bg=UI_COLORS["card"])
        col_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        tk.Label(col_path, text="保存路径",
                 bg=UI_COLORS["card"], fg=UI_COLORS["text_secondary"],
                 font=(UI_FONT, 9)).pack(anchor=tk.W)
        path_row = tk.Frame(col_path, bg=UI_COLORS["card"])
        path_row.pack(fill=tk.X)
        path_entry = tk.Entry(path_row, textvariable=self.output_var, **_entry_style)
        path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=6)
        browse_btn = tk.Button(
            path_row, text="浏览", command=self._browse_output,
            bg=UI_COLORS["card_alt"], fg=UI_COLORS["text_secondary"],
            activebackground="#EEF2F7", relief=tk.FLAT, bd=0,
            font=(UI_FONT, 9), padx=14, pady=4, cursor="hand2",
        )
        browse_btn.pack(side=tk.LEFT, padx=(6, 0))

        # XPath
        col_xpath = tk.Frame(row1, bg=UI_COLORS["card"])
        col_xpath.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        tk.Label(col_xpath, text="XPath",
                 bg=UI_COLORS["card"], fg=UI_COLORS["text_secondary"],
                 font=(UI_FONT, 9)).pack(anchor=tk.W)
        xpath_entry = tk.Entry(col_xpath, textvariable=self.xpath_var,
                               **{**_entry_style, "font": (MONO_FONT, 10)})
        xpath_entry.pack(fill=tk.X, ipady=6)

        # 最大页数
        col_pages = tk.Frame(row1, bg=UI_COLORS["card"])
        col_pages.pack(side=tk.LEFT, fill=tk.X, expand=False, padx=(0, 0))
        tk.Label(col_pages, text="最大页数",
                 bg=UI_COLORS["card"], fg=UI_COLORS["text_secondary"],
                 font=(UI_FONT, 9)).pack(anchor=tk.W)
        pages_entry = tk.Entry(col_pages, textvariable=self.max_pages_var,
                               width=10, **_entry_style)
        pages_entry.pack(fill=tk.X, ipady=6)

        # 第二行: 按钮
        row2 = tk.Frame(ctrl_inner, bg=UI_COLORS["card"])
        row2.pack(fill=tk.X)

        self.stop_btn = tk.Button(
            row2, text="停止", command=self._stop,
            bg=UI_COLORS["card"], fg=UI_COLORS["danger"],
            activebackground=UI_COLORS["danger_light"],
            relief=tk.FLAT, bd=1, font=(UI_FONT, 10, "bold"),
            padx=20, pady=8, state=tk.DISABLED, cursor="hand2",
            highlightthickness=1, highlightbackground="#FECACA",
            highlightcolor="#FECACA",
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.start_btn = tk.Button(
            row2, text="开始采集", command=self._start,
            bg=UI_COLORS["primary"], fg="#FFFFFF",
            activebackground=UI_COLORS["primary_hover"],
            activeforeground="#FFFFFF",
            relief=tk.FLAT, bd=0, font=(UI_FONT, 10, "bold"),
            padx=28, pady=8, cursor="hand2",
        )
        self.start_btn.pack(side=tk.LEFT)

        # ── 分类列表 ──
        pages_card = tk.Frame(main, bg=UI_COLORS["card"],
                              highlightbackground=UI_COLORS["border"],
                              highlightthickness=1, bd=0)
        pages_card.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        pages_inner = tk.Frame(pages_card, bg=UI_COLORS["card"])
        pages_inner.pack(fill=tk.BOTH, expand=True, padx=16, pady=12)

        tk.Label(
            pages_inner, text="分类列表（每行: 标题, URL）",
            bg=UI_COLORS["card"], fg=UI_COLORS["text"],
            font=(UI_FONT, 11, "bold"),
        ).pack(anchor=tk.W)
        tk.Label(
            pages_inner,
            text='示例: Loungewear > Pajamas, https://example.com/collections/pajamas',
            bg=UI_COLORS["card"], fg=UI_COLORS["text_muted"],
            font=(UI_FONT, 9),
        ).pack(anchor=tk.W, pady=(2, 8))

        self.pages_text = scrolledtext.ScrolledText(
            pages_inner, height=8, wrap=tk.NONE,
            bg=UI_COLORS["card_alt"], fg=UI_COLORS["text"],
            insertbackground=UI_COLORS["primary"],
            font=(MONO_FONT, 9), relief=tk.FLAT, bd=0,
            highlightthickness=1, highlightbackground=UI_COLORS["border"],
            highlightcolor=UI_COLORS["primary"],
            padx=10, pady=8,
        )
        self.pages_text.pack(fill=tk.BOTH, expand=True)

        # ── 日志区 ──
        log_card = tk.Frame(main, bg=UI_COLORS["log_background"],
                            highlightbackground="#1E293B",
                            highlightthickness=1, bd=0)
        log_card.pack(fill=tk.BOTH, expand=True)

        log_header = tk.Frame(log_card, bg=UI_COLORS["log_panel"], height=36)
        log_header.pack(fill=tk.X)
        log_header.pack_propagate(False)
        tk.Label(
            log_header, text="运行日志",
            bg=UI_COLORS["log_panel"], fg="#F8FAFC",
            font=(UI_FONT, 10, "bold"),
        ).pack(side=tk.LEFT, padx=14, pady=8)
        clear_btn = tk.Button(
            log_header, text="清空", command=self._clear_log,
            bg="#1E293B", fg="#CBD5E1", activebackground="#334155",
            activeforeground="#FFFFFF", relief=tk.FLAT, bd=0,
            font=(UI_FONT, 8), padx=10, pady=4, cursor="hand2",
        )
        clear_btn.pack(side=tk.RIGHT, padx=10, pady=6)

        self.log_text = scrolledtext.ScrolledText(
            log_card, state=tk.DISABLED,
            bg=UI_COLORS["log_background"], fg=UI_COLORS["log_text"],
            insertbackground="#FFFFFF", selectbackground="#1D4ED8",
            selectforeground="#FFFFFF",
            font=(MONO_FONT, 9), relief=tk.FLAT, bd=0,
            padx=12, pady=10, wrap=tk.WORD,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _browse_output(self):
        path = filedialog.askdirectory(title="选择保存目录")
        if path:
            self.output_var.set(path)

    def _clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _parse_pages(self):
        """解析分类列表文本框，返回 [(title, url), ...]"""
        raw = self.pages_text.get("1.0", tk.END).strip()
        if not raw:
            return []
        pages = []
        for line in raw.splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue

            # 提取 URL（从 http 开始到下一个引号/括号/逗号截止）
            url_match = re.search(r'(https?://[^\s"\')\]]+)', line)
            if not url_match:
                continue
            url = url_match.group(1)

            # 提取标题：URL 前面的文本，去掉括号、引号、逗号
            title = ""
            url_start = line.find(url)
            if url_start > 0:
                title_part = line[:url_start].strip(' \t,("\'')
                title = re.sub(r'[",)\]]+$', '', title_part).strip()

            if not title:
                title = urlparse(url).path.split("/")[-1]

            pages.append((title, url))
        return pages

    def _start(self):
        pages = self._parse_pages()
        if not pages:
            messagebox.showerror("错误", "请先填写分类列表")
            return

        output_path = self.output_var.get().strip()
        if not output_path:
            output_path = os.path.join(
                os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                "output",
            )
            self.output_var.set(output_path)

        max_pages_raw = self.max_pages_var.get().strip()
        try:
            max_pages = int(max_pages_raw) if max_pages_raw else 0
        except ValueError:
            max_pages = 0

        xpath = self.xpath_var.get().strip()

        self.is_running = True
        self.stop_event.clear()
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self._clear_log()

        thread = threading.Thread(
            target=self._run_task,
            args=(pages, output_path, max_pages, xpath),
            daemon=True,
        )
        thread.start()

    def _stop(self):
        if self.is_running:
            self.stop_event.set()
            self.stop_btn.config(state=tk.DISABLED)
            print("\n[已停止] 等待当前页完成后退出...")

    def _run_task(self, pages, output_path, max_pages, xpath=""):
        try:
            print(f"共 {len(pages)} 个分类")
            if max_pages:
                print(f"最大页数: {max_pages}")
            else:
                print(f"最大页数: 不限制")
            if xpath:
                print(f"XPath: {xpath}")
            print(f"保存目录: {output_path}")
            print()

            chrome = create_browser()
            tab = chrome.latest_tab
            all_results = []
            seen_all_links = set()

            try:
                for i, (title, base_url) in enumerate(pages, 1):
                    if self.stop_event.is_set():
                        break

                    try:
                        print(f"[{i}/{len(pages)}] {title}")

                        page_links = []
                        seen_in_category = set()
                        page_num = 1

                        while True:
                            if self.stop_event.is_set():
                                break
                            if max_pages and page_num > max_pages:
                                print(f"  已达到最大页数 {max_pages}")
                                break

                            if page_num == 1:
                                url = base_url
                            else:
                                sep = "&" if "?" in base_url else "?"
                                url = f"{base_url}{sep}page={page_num}"

                            print(f"  第 {page_num} 页: {url}")
                            try:
                                tab.get(url)
                                links, infinite_scroll = collect_page_links(
                                    tab,
                                    stop_event=self.stop_event,
                                    xpath=xpath,
                                )
                            except ContextLostError:
                                print("  页面刷新，等待恢复...")
                                time.sleep(3)
                                try:
                                    tab.get(url)
                                    links, infinite_scroll = collect_page_links(
                                        tab,
                                        stop_event=self.stop_event,
                                        xpath=xpath,
                                    )
                                except ContextLostError:
                                    print("  页面未恢复，跳过该页")
                                    break

                            new_links = [l for l in links if l not in seen_in_category]

                            if not new_links:
                                if page_num == 1:
                                    retried = False
                                    for retry in range(1, MAX_RETRIES + 1):
                                        print(f"  首页无链接，第 {retry}/{MAX_RETRIES} 次重试...")
                                        time.sleep(RETRY_DELAY)
                                        tab.get(url)
                                        links, infinite_scroll = collect_page_links(
                                            tab,
                                            stop_event=self.stop_event,
                                            xpath=xpath,
                                        )
                                        new_links = [
                                            l for l in links
                                            if l not in seen_in_category
                                        ]
                                        if new_links:
                                            print(f"  重试成功: {len(new_links)} 条")
                                            retried = True
                                            break
                                    if not retried:
                                        print("  跳过该分类")
                                        break
                                else:
                                    print("  无新链接，该分类采集完毕")
                                    break

                            for link in new_links:
                                seen_in_category.add(link)
                                page_links.append(link)

                            print(f"  新增 {len(new_links)} 条（累计 {len(page_links)} 条）")

                            if infinite_scroll:
                                print("  无限滚动分类采集完毕")
                                break

                            page_num += 1

                        new_for_all = 0
                        for link in page_links:
                            if link in seen_all_links:
                                continue
                            seen_all_links.add(link)
                            all_results.append({"title": title, "link": link})
                            new_for_all += 1

                        skipped = len(page_links) - new_for_all
                        print(f"  共采集 {len(page_links)} 条（新增 {new_for_all} 条，跨分类去重跳过 {skipped} 条）")
                        print()

                    except Exception as e:
                        print(f"  分类 [{title}] 出错: {e}")
                        print()
                        continue

            finally:
                chrome.quit()

            print(f"采集完成: 共 {len(all_results)} 条链接")

            os.makedirs(output_path, exist_ok=True)

            # 记录 域名→分类标题→分类URL 的映射（JSON 索引，可追加）
            update_link_index(output_path, pages)

            if all_results:
                # 按域名分组
                domain_groups = {}
                for row in all_results:
                    domain = urlparse(row["link"]).netloc
                    if domain not in domain_groups:
                        domain_groups[domain] = []
                    domain_groups[domain].append(row)

                for domain, rows in domain_groups.items():
                    file_name = f"{domain}.xlsx"
                    file_path = os.path.join(output_path, file_name)

                    # 合并已存在的同名文件，按 link 去重，避免覆盖历史采集结果
                    existing = load_existing_domain_rows(file_path)
                    merged = dict(existing)
                    added = 0
                    for row in rows:
                        link = row["link"]
                        if link not in merged:
                            merged[link] = row.get("title", "")
                            added += 1

                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.append(["title", "link"])
                    for link, title in merged.items():
                        ws.append([title, link])
                    wb.save(file_path)
                    print(
                        f"已保存: {file_path}（已有 {len(existing)} 条 + 新增 {added} 条"
                        f" = 合并 {len(merged)} 条）"
                    )

                print(f"共 {len(domain_groups)} 个域名文件")
            else:
                print("未采集到任何链接")

        except Exception as e:
            print(f"发生异常: {e}")
            import traceback
            traceback.print_exc()

        finally:
            self.is_running = False
            self.root.after(0, self._reset_buttons)

    def _reset_buttons(self):
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)


if __name__ == "__main__":
    root = tk.Tk()
    app = LinkCollectorApp(root)
    root.mainloop()
