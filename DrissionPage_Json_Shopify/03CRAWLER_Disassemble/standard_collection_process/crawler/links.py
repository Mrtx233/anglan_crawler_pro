from __future__ import annotations

from DrissionPage.errors import ContextLostError
from urllib.parse import urlparse
import json
import re
import time
from ..config import PAGE_LOAD_TIMEOUT


def extract_product_links(html_text, base_url):
    """普通模式只从 JSON handle 构造商品链接，不回退到 DOM 链接。"""
    parsed = urlparse(base_url)
    base = f"{parsed.scheme}://{parsed.netloc}"
    handles = dict.fromkeys(
        handle.strip()
        for handle in re.findall(r'"handle"\s*:\s*"([^"\\]+)"', html_text)
        if handle.strip()
    )
    return [f"{base}/products/{handle}" for handle in handles]


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
    """文档加载完成后直接提取；指定 XPath 时只使用 XPath。"""
    if stop_event is not None and stop_event.is_set():
        return []
    if not tab.wait.doc_loaded(timeout=PAGE_LOAD_TIMEOUT, raise_err=False):
        raise TimeoutError(f"页面加载超时: {tab.url}")
    if stop_event is not None and stop_event.is_set():
        return []
    if xpath:
        return extract_links_by_xpath(tab, xpath)
    html_text = tab.run_js("return document.documentElement.outerHTML;") or ""
    return extract_product_links(html_text, tab.url)
