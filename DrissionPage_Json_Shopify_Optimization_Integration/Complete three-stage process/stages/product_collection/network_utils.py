# ======================== 网络与 URL 层 ========================
#
# 职责: Shopify JSON URL 构造、目标分组去重、汇率获取与货币换算、
#       通过 DrissionPage 监听网络包获取 JSON。
# 依赖: config（汇率缓存）。仅此一项。
# 禁止: 导入 tkinter、product_parser。
#       （product_parser 会导入本模块取 convert_price_to_usd，
#        反向导入将形成循环）
#
# 汇率缓存位置说明:
#   _exchange_rates_cache 定义在 config 而非本模块，
#   因为 product_parser 也要读它，放这里会造成循环导入。
#   本模块通过 config._exchange_rates_cache = ... 写入。
# ===============================================================

import json
import ssl
import time
import urllib.request
from urllib.parse import urlparse, urlunparse

import certifi

from stages.product_collection import config

# ════════════════════════════════════════════════════════════
# 核心业务逻辑 (保持原有逻辑不变)
# ════════════════════════════════════════════════════════════
def to_shopify_json_url(url):
    """把商品详情页 URL 转成 Shopify 的 .json 接口地址。

    去掉末尾斜杠后追加 .json，并丢弃 query 与 fragment
    （Shopify 的 json 接口不接受这些参数）。
    """
    parsed = urlparse(url)
    path = parsed.path.rstrip("/")
    if not path.endswith(".json"):
        path = f"{path}.json"
    return urlunparse((parsed.scheme, parsed.netloc, path, "", "", ""))

def group_targets_by_json_url(targets):
    """把 (title, url) 列表按 json 接口地址分组去重。

    返回 (分组字典, 重复商品数)。分组字典的键是 json URL，值是
    {title, original_url}，其中 original_url 已剥掉 query 与 fragment，
    用于后续写入 link-href 并做断点续采比对。
    完全相同的商品（同一 json URL）只保留首次出现的那条。
    """
    json_url_groups = {}
    duplicate_product_count = 0
    for title, url in targets:
        original_url = str(url).strip()
        if not original_url:
            continue
        json_url = to_shopify_json_url(original_url)
        if json_url in json_url_groups:
            duplicate_product_count += 1
            continue
        title_text = str(title).strip() if title is not None else ""
        json_url_groups[json_url] = {
            "title": title_text,
            "original_url": urlunparse(urlparse(original_url)._replace(query="", fragment="")),
        }
    return json_url_groups, duplicate_product_count

def fetch_exchange_rates():
    """获取 USD 基准汇率并缓存到 config；已有缓存时直接返回。"""
    if config._exchange_rates_cache:
        return config._exchange_rates_cache
    try:
        url = "https://open.er-api.com/v6/latest/USD"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        # 保留系统/自定义 CA，并补充 certifi，兼容 macOS Python 缺失默认 CA 的环境。
        context = ssl.create_default_context()
        context.load_verify_locations(cafile=certifi.where())
        with urllib.request.urlopen(req, timeout=10, context=context) as response:
            data = json.loads(response.read().decode())
        config._exchange_rates_cache = data.get("rates", {})
        print(f"[汇率] 成功获取 {config._exchange_rates_cache.get('USD', 1)} 相关汇率")
    except Exception as e:
        print(f"[汇率] 获取失败: {e}")
    return config._exchange_rates_cache

def convert_price_to_usd(price_str, currency, rates=None):
    """按汇率把价格换算成 USD 并保留两位小数。

    任一前提缺失（无价格、无币种、无汇率表、未知币种、价格非数字）
    都原样返回入参字符串——宁可保留原值，也不让采集因汇率问题中断。
    """
    if not price_str or not currency:
        return price_str
    if rates is None:
        rates = config._exchange_rates_cache
    if not rates:
        return price_str
    currency = str(currency).upper().strip()
    if currency == "USD":
        return price_str
    rate = rates.get(currency)
    if not rate or rate == 0:
        return price_str
    try:
        amount = float(str(price_str).strip())
    except (ValueError, TypeError):
        return price_str
    usd_amount = amount / rate
    return f"{usd_amount:.2f}"

def fetch_json(tab, json_url, stop_event=None):
    """获取 Shopify JSON；网络包等待按 1 秒切片，便于停止事件及时生效。"""
    tab.listen.start(json_url)
    tab.get(json_url)

    packet = None
    deadline = time.monotonic() + 30
    while time.monotonic() < deadline:
        if stop_event is not None and stop_event.is_set():
            return 0, None
        remaining = deadline - time.monotonic()
        packet = tab.listen.wait(timeout=min(1.0, max(0.05, remaining)))
        if packet:
            break

    status = packet.response.status if packet else 0

    if status == 200:
        if stop_event is not None and stop_event.is_set():
            return 0, None
        text = tab.run_js("return document.body.innerText")
        if not text:
            text = tab.run_js("return document.body.textContent")
        try:
            return status, json.loads(text)
        except json.JSONDecodeError:
            pre_text = tab.run_js("return document.querySelector('pre')?.innerText || ''")
            if pre_text:
                try:
                    return status, json.loads(pre_text)
                except json.JSONDecodeError:
                    pass
            return status, None
    return status, None
