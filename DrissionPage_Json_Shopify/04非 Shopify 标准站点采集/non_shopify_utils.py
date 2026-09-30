# -*- coding: utf-8 -*-
"""非 Shopify 站点采集脚本共用的纯数据处理工具。

本模块不依赖 Shopify 主采集器、GUI、浏览器对象或项目目录结构。
价格语义与现有主管线保持一致：无小数点的数值按 cents 处理。
"""

import json
import re
import urllib.request
from html import escape
from html.parser import HTMLParser
from urllib.parse import urlparse, urlunparse


_exchange_rates_cache = {}


def fetch_exchange_rates():
    """返回以 USD 为基准的汇率映射；请求失败时返回空字典。"""
    global _exchange_rates_cache
    if _exchange_rates_cache:
        return _exchange_rates_cache

    try:
        request = urllib.request.Request(
            "https://open.er-api.com/v6/latest/USD",
            headers={"User-Agent": "Mozilla/5.0"},
        )
        with urllib.request.urlopen(request, timeout=10) as response:
            data = json.loads(response.read().decode("utf-8"))
        rates = data.get("rates", {})
        _exchange_rates_cache = rates if isinstance(rates, dict) else {}
        print(f"[汇率] 成功获取 {_exchange_rates_cache.get('USD', 1)} 相关汇率")
    except Exception as exc:
        _exchange_rates_cache = {}
        print(f"[汇率] 获取失败: {exc}")
    return _exchange_rates_cache


def convert_price_to_usd(price_str, currency, rates=None):
    """使用 USD 基准汇率表把主币单位价格转换为 USD 字符串。"""
    if not price_str or not currency:
        return price_str
    if rates is None:
        rates = _exchange_rates_cache
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
    return f"{amount / rate:.2f}"


def format_price(value, rates=None, currency=None):
    """格式化价格，并保持主管线的 cents/主币单位兼容语义。"""
    if value in (None, ""):
        return ""
    raw = str(value).strip()
    if not raw:
        return ""
    if rates and currency:
        raw = convert_price_to_usd(raw, currency, rates)

    if "." in raw:
        try:
            amount = float(raw)
        except ValueError:
            return raw
    else:
        try:
            amount = int(raw) / 100
        except ValueError:
            return raw
    return str(int(amount)) if amount.is_integer() else f"{amount:.2f}".rstrip("0").rstrip(".")


def format_image_url(url, size="600x600", use_query_marker=False):
    """添加图片尺寸标记，并完整保留原有 query/fragment。

    默认把尺寸写入文件名。对于不支持改写文件名的 CDN，可使用
    ``use_query_marker=True`` 保留原路径并追加 ``_600x600=1``。
    """
    if not url:
        return ""
    parsed = urlparse(str(url))
    if use_query_marker:
        marker_name = f"_{size}"
        query_parts = [part for part in parsed.query.split("&") if part]
        if not any(part.split("=", 1)[0] == marker_name for part in query_parts):
            query_parts.append(f"{marker_name}=1")
        return urlunparse((
            parsed.scheme,
            parsed.netloc,
            parsed.path,
            parsed.params,
            "&".join(query_parts),
            parsed.fragment,
        ))

    path = parsed.path
    dot_index = path.rfind(".")
    slash_index = path.rfind("/")
    if dot_index > slash_index:
        stem = path[:dot_index]
        suffix = path[dot_index:]
        dimension_re = re.compile(r"_\d+x\d+$", re.IGNORECASE)
        if dimension_re.search(stem):
            stem = dimension_re.sub(f"_{size}", stem)
        elif not stem.lower().endswith(f"_{size}".lower()):
            stem = f"{stem}_{size}"
        path = f"{stem}{suffix}"
    return urlunparse((
        parsed.scheme,
        parsed.netloc,
        path,
        parsed.params,
        parsed.query,
        parsed.fragment,
    ))


class CleanBodyHtmlParser(HTMLParser):
    """移除链接、图片及属性，保留主管线所需的基础 HTML 结构。"""

    void_tags = {
        "area", "base", "br", "col", "embed", "hr", "img", "input",
        "link", "meta", "param", "source", "track", "wbr",
    }

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []

    def handle_starttag(self, tag, attrs):
        if tag in ("a", "img"):
            return
        self.parts.append(f"<{tag}>")

    def handle_endtag(self, tag):
        if tag == "a" or tag in self.void_tags:
            return
        self.parts.append(f"</{tag}>")

    def handle_startendtag(self, tag, attrs):
        if tag in ("a", "img"):
            return
        self.parts.append(f"<{tag}>")

    def handle_data(self, data):
        data = "".join(character for character in data if ord(character) <= 0xFFFF)
        self.parts.append(escape(data, quote=False))

    def get_html(self):
        result = "".join(self.parts).strip()
        empty_tag_re = re.compile(r"<([a-zA-Z][a-zA-Z0-9]*)\b[^>]*>\s*</\1>")
        while True:
            result, count = empty_tag_re.subn("", result)
            if count == 0:
                return result


def clean_body_html(body_html):
    """清洗商品详情 HTML；不读取或修改任何外部状态。"""
    if not body_html:
        return ""
    parser = CleanBodyHtmlParser()
    parser.feed(str(body_html))
    parser.close()
    return parser.get_html()


__all__ = [
    "clean_body_html",
    "convert_price_to_usd",
    "fetch_exchange_rates",
    "format_image_url",
    "format_price",
]
