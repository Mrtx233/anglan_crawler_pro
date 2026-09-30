import csv
import hashlib
import json
import os
import random
import re
import string
import sys
import time
import urllib.request
from html import escape
from html.parser import HTMLParser
from itertools import product as itertools_product
from pathlib import Path
from urllib.parse import parse_qs, urlparse, urlunparse
import threading

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import pandas as pd
import openpyxl
from DrissionPage import Chromium, ChromiumOptions

# ════════════════════════════════════════════════════════════
# 全局采集配置
# ════════════════════════════════════════════════════════════
DELAY = 0.5
MAX_RETRIES = 5
RETRY_BASE_DELAY = 30
SAVE_EVERY = 20
SKIP_POSITIONS = None
KEEP_POSITIONS = None
SKIP_OPTIONS = ["ships from"]  # 屏蔽"发货地"选项，它不属于商品变体

CSV_FIELDS = [
    "title", "name", "price1", "price2",
    "styles1", "styles2", "styles3",
    "src_links", "link-href", "details",
]


# ════════════════════════════════════════════════════════════
# 核心业务逻辑 (保持原有逻辑不变)
# ════════════════════════════════════════════════════════════
def to_shopify_json_url(url):
    parsed = urlparse(url)
    path = parsed.path.rstrip("/")
    if not path.endswith(".json"):
        path = f"{path}.json"
    return urlunparse((parsed.scheme, parsed.netloc, path, "", "", ""))


def group_targets_by_json_url(targets):
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


_exchange_rates_cache = {}


def fetch_exchange_rates():
    global _exchange_rates_cache
    if _exchange_rates_cache:
        return _exchange_rates_cache
    try:
        url = "https://open.er-api.com/v6/latest/USD"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as response:
            data = json.loads(response.read().decode())
        _exchange_rates_cache = data.get("rates", {})
        print(f"[汇率] 成功获取 {_exchange_rates_cache.get('USD', 1)} 相关汇率")
    except Exception as e:
        print(f"[汇率] 获取失败: {e}")
    return _exchange_rates_cache


def convert_price_to_usd(price_str, currency, rates=None):
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
    usd_amount = amount / rate
    return f"{usd_amount:.2f}"


def find_fallback_variant(product):
    variants = product.get("variants", [])
    for variant in variants:
        if variant.get("position") == 1:
            return variant
    return variants[0] if variants else {}


def format_combo_value(value):
    return "" if value is None else str(value)


SIZE_VALUES = {
    "xs", "s", "m", "l", "xl", "xxl", "xxxl", "xxxxl",
    "2xl", "3xl", "4xl", "5xl",
    "xs/s", "s/m", "m/l", "l/xl",
    "one size", "os", "free size",
}


def _is_size_option(option):
    """判断 option 是否为尺码选项。名称含 'size' 或 values 都是尺码值。"""
    name = (option.get("name") or "").lower().strip()
    if "size" in name:
        return True
    values = [str(v).lower().strip() for v in (option.get("values") or [])]
    if values and all(v in SIZE_VALUES for v in values):
        return True
    return False


def get_sorted_options(product, skip_options=None):
    options = sorted(
        product.get("options", []),
        key=lambda option: option.get("position") or 0,
    )
    if skip_options:
        skip_lower = [s.lower() for s in skip_options]
        options = [
            opt for opt in options
            if not any(s in (opt.get("name") or "").lower() for s in skip_lower)
        ]
    multi_value = [opt for opt in options if len(opt.get("values") or []) > 1]
    # 统一尺码名称为 "Size"
    for opt in multi_value:
        if _is_size_option(opt):
            opt["name"] = "Size"
    # "Type" 是 Shopify 保留属性，作为 Option 名导入会报错，统一改名为 "Style"
    for opt in multi_value:
        if (opt.get("name") or "").strip().lower() == "type":
            opt["name"] = "Style"
    color_idx = next(
        (i for i, opt in enumerate(multi_value)
         if (opt.get("name") or "").lower() == "color"),
        None,
    )
    if color_idx is not None and color_idx > 0:
        color_opt = multi_value.pop(color_idx)
        multi_value.insert(0, color_opt)
    return multi_value


def build_option_combos(options):
    value_groups = [
        [format_combo_value(value) for value in option.get("values", [])]
        for option in options
    ]
    if not value_groups:
        return [()]
    return list(itertools_product(*value_groups))


def _option_field(options, combo_index):
    if combo_index < len(options):
        return f"option{options[combo_index].get('position', combo_index + 1)}"
    return f"option{combo_index + 1}"


def _option_field_by_position(options, position):
    """根据 option 的实际 position 返回 variant 字段名。"""
    return f"option{position}"


def variant_matches_combo(variant, combo, options):
    for index, value in enumerate(combo):
        field = _option_field(options, index)
        if format_combo_value(variant.get(field)) != format_combo_value(value):
            return False
    return True


def find_variant_by_combo(product, combo, options):
    for variant in product.get("variants", []):
        if variant_matches_combo(variant, combo, options):
            return variant
    return {}


def find_variant_by_first_option(product, first_option_value, options):
    target = format_combo_value(first_option_value)
    field = _option_field(options, 0)
    for variant in product.get("variants", []):
        if format_combo_value(variant.get(field)) == target:
            return variant
    return {}


def format_price(value, rates=None, currency=None):
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


def _is_zero_price(value):
    try:
        return float(str(value).strip()) == 0
    except (ValueError, TypeError):
        return False


def format_price2(variant, rates=None):
    currency = variant.get("price_currency", "")
    compare_at_price = variant.get("compare_at_price")
    if compare_at_price and str(compare_at_price).strip() and not _is_zero_price(compare_at_price):
        return format_price(compare_at_price, rates=rates, currency=currency)
    return format_price(variant.get("price", ""), rates=rates, currency=currency)


def format_image_url(url, size="600x600"):
    if not url:
        return ""
    parsed = urlparse(url)
    path = parsed.path
    dot_index = path.rfind(".")
    slash_index = path.rfind("/")
    if dot_index > slash_index:
        path = f"{path[:dot_index]}_{size}{path[dot_index:]}"
    return urlunparse((
        parsed.scheme, parsed.netloc, path,
        parsed.params, parsed.query, parsed.fragment,
    ))


def _resolve_positions(images, positions):
    """将负数解析为实际位置（-1 = 最后一张），返回实际 position 值的 set。"""
    if not positions:
        return set()
    all_positions = sorted(set(img.get("position") or 0 for img in images))
    resolved = set()
    for sp in positions:
        if sp < 0:
            idx = -sp
            if idx <= len(all_positions):
                resolved.add(all_positions[-idx])
        else:
            resolved.add(sp)
    return resolved


def _filter_images(images, keep_positions=None, skip_positions=None):
    """过滤图片列表。keep 优先于 skip。"""
    if keep_positions is not None:
        resolved = _resolve_positions(images, keep_positions)
        return [img for img in images if (img.get("position") or 0) in resolved]
    if skip_positions is not None:
        resolved = _resolve_positions(images, skip_positions)
        return [img for img in images if (img.get("position") or 0) not in resolved]
    return images


def build_images_by_id(product, keep_positions=None, skip_positions=None):
    images = _filter_images(product.get("images", []), keep_positions=keep_positions, skip_positions=skip_positions)
    return {
        image.get("id"): format_image_url(image.get("src", ""))
        for image in images
    }


def build_first_image_srcs(product, variant_id=None, limit=None, keep_positions=None, skip_positions=None):
    images = sorted(
        product.get("images", []),
        key=lambda image: image.get("position") or 0,
    )
    images = _filter_images(images, keep_positions=keep_positions, skip_positions=skip_positions)
    if variant_id is not None:
        variant_images = [
            image for image in images
            if variant_id in (image.get("variant_ids") or [])
        ]
        if variant_images:
            images = variant_images
    selected_images = images if limit is None else images[:limit]
    srcs = [
        format_image_url(image.get("src", ""))
        for image in selected_images
        if image.get("src")
    ]
    return "#".join(srcs)


def build_images_by_color_option(product, options, keep_positions=None, skip_positions=None):
    """按主选项建立 {选项值: 图片URL} 映射。

    优先用名为 color 的选项（get_sorted_options 已将其排到最前）；
    没有 color 时退回第一个选项（如 Device/Size），让单选项商品的
    每个 variant 也能带上自己的 image_id 对应图。
    """
    images_by_id = build_images_by_id(product, keep_positions=keep_positions, skip_positions=skip_positions)
    color_opt = next((opt for opt in options if (opt.get("name") or "").lower() == "color"), None)
    if not color_opt:
        color_opt = options[0] if options else None
    if not color_opt:
        return {}
    field = _option_field_by_position(options, color_opt.get("position", 1))
    result = {}
    for variant in product.get("variants", []):
        opt_value = format_combo_value(variant.get(field))
        if opt_value and opt_value not in result:
            image_id = variant.get("image_id")
            if image_id and image_id in images_by_id:
                result[opt_value] = images_by_id[image_id]
    return result


# ── styles 文本转义 ──────────────────────────────────────────
# styles 用 & # @ 作字段分隔符；选项值/选项名若含这些字符（如 "Tie & Square"）
# 会被误切成多个字段导致乱码。写入时对保留字符做反斜杠转义，解析时按层还原。
_STYLE_RESERVED = r"\&#@"


def _style_escape(value):
    """转义选项值/选项名里的保留分隔符（\\ & # @）；不含则原样返回。"""
    if value is None:
        return ""
    text = str(value)
    if not any(ch in text for ch in _STYLE_RESERVED):
        return text
    return (text.replace("\\", "\\\\")
                .replace("&", "\\&")
                .replace("#", "\\#")
                .replace("@", "\\@"))


def _style_split(text, sep):
    """按未转义的 sep 拆分，保留转义序列本身（供后续按层还原）。"""
    parts = []
    buf = []
    i, n = 0, len(text)
    while i < n:
        ch = text[i]
        if ch == "\\" and i + 1 < n and text[i + 1] in _STYLE_RESERVED:
            buf.append(text[i:i + 2])
            i += 2
            continue
        if ch == sep:
            parts.append("".join(buf))
            buf = []
            i += 1
            continue
        buf.append(ch)
        i += 1
    parts.append("".join(buf))
    return parts


def _style_split_first(text, sep):
    """在第一个未转义的 sep 处切一刀，返回 (前段, 后段, 是否找到)。"""
    i, n = 0, len(text)
    while i < n:
        ch = text[i]
        if ch == "\\" and i + 1 < n and text[i + 1] in _STYLE_RESERVED:
            i += 2
            continue
        if ch == sep:
            return text[:i], text[i + 1:], True
        i += 1
    return text, "", False


def _style_unescape(text):
    """还原 _style_escape 转义过的字符。"""
    if "\\" not in text:
        return text
    out = []
    i, n = 0, len(text)
    while i < n:
        ch = text[i]
        if ch == "\\" and i + 1 < n and text[i + 1] in _STYLE_RESERVED:
            out.append(text[i + 1])
            i += 2
            continue
        out.append(ch)
        i += 1
    return "".join(out)


def build_segment(combo, options, variant, images_by_color_option, rates=None):
    segment_parts = []
    if combo:
        segment_parts.append(_style_escape(format_combo_value(combo[0])))
        for index, value in enumerate(combo[1:], start=1):
            option_name = format_combo_value(options[index].get("name", ""))
            segment_parts.extend([
                _style_escape(option_name),
                _style_escape(format_combo_value(value)),
            ])
    currency = variant.get("price_currency", "")
    price1 = format_price(variant.get("price"), rates=rates, currency=currency)
    price2 = format_price2(variant, rates=rates)
    color_value = format_combo_value(combo[0]) if combo else ""
    image_src = images_by_color_option.get(color_value, "")
    segment = f"{'&'.join(segment_parts)}${price1}${price2}"
    if image_src:
        segment = f"{segment}@{image_src}"
    return segment


def build_variant_combo(product, keep_positions=None, skip_positions=None, skip_options=None, rates=None):
    options = get_sorted_options(product, skip_options=skip_options)
    combos = build_option_combos(options)
    fallback_variant = find_fallback_variant(product)
    images_by_color_opt = build_images_by_color_option(product, options, keep_positions=keep_positions, skip_positions=skip_positions)
    segments = []
    for combo in combos:
        variant = find_variant_by_combo(product, combo, options)
        if not variant:
            variant = find_variant_by_first_option(product, combo[0], options) if combo else {}
        if not variant:
            variant = fallback_variant
        segments.append(build_segment(combo, options, variant, images_by_color_opt, rates=rates))
    first_option_name = format_combo_value(options[0].get("name", "")) if options else ""
    if first_option_name:
        return f"{_style_escape(first_option_name)}#{'#'.join(segments)}"
    return "#".join(segments)


class CleanBodyHtmlParser(HTMLParser):
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
        data = "".join(ch for ch in data if ord(ch) <= 0xFFFF)
        self.parts.append(escape(data, quote=False))

    def get_html(self):
        html = "".join(self.parts).strip()
        empty_tag_re = re.compile(r"<([a-zA-Z][a-zA-Z0-9]*)\b[^>]*>\s*</\1>")
        while True:
            cleaned, n = empty_tag_re.subn("", html)
            if n == 0:
                break
            html = cleaned
        return html


def clean_body_html(body_html):
    if not body_html:
        return ""
    parser = CleanBodyHtmlParser()
    parser.feed(str(body_html))
    parser.close()
    return parser.get_html()


def create_browser():
    co = ChromiumOptions()
    co.incognito(on_off=True)
    # 代理设置，如有需要自行取消或修改
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


def read_targets_from_excel(filepath):
    wb = openpyxl.load_workbook(filepath, read_only=True)
    ws = wb.active
    targets = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        title = str(row[0]).strip() if row and row[0] else ""
        url = str(row[1]).strip() if len(row) > 1 and row[1] else ""
        if url:
            targets.append((title, url))
    wb.close()
    return targets


def fetch_json(tab, json_url):
    tab.listen.start(json_url)
    tab.get(json_url)
    packet = tab.listen.wait(timeout=30)
    status = packet.response.status if packet else 0

    if status == 200:
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
    else:
        return status, None


def parse_product(data, table_title, original_url):
    product = data.get("product", {})
    variant = find_fallback_variant(product)
    options = get_sorted_options(product, skip_options=SKIP_OPTIONS)
    variant_image_urls = [
        url for url in build_images_by_color_option(
            product, options, keep_positions=KEEP_POSITIONS, skip_positions=SKIP_POSITIONS
        ).values() if url
    ]
    src_links = [
        src for src in build_first_image_srcs(
            product, keep_positions=KEEP_POSITIONS, skip_positions=SKIP_POSITIONS
        ).split("#") if src
    ]
    for url in variant_image_urls:
        if url not in src_links:
            src_links.append(url)
    return {
        "title": table_title,
        "name": product.get("title", ""),
        "price1": format_price(
            variant.get("price", ""),
            rates=_exchange_rates_cache,
            currency=variant.get("price_currency", ""),
        ),
        "price2": format_price2(variant, rates=_exchange_rates_cache),
        "styles1": build_variant_combo(
            product,
            keep_positions=KEEP_POSITIONS,
            skip_positions=SKIP_POSITIONS,
            skip_options=SKIP_OPTIONS,
            rates=_exchange_rates_cache,
        ),
        "styles2": "",
        "styles3": "",
        "src_links": "#".join(src_links),
        "link-href": original_url,
        "details": clean_body_html(product.get("body_html", "")),
    }


def save_xlsx(rows, output_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(CSV_FIELDS)
    for row in rows:
        ws.append([row.get(f, "") for f in CSV_FIELDS])
    wb.save(output_path)


def load_xlsx(filepath):
    wb = openpyxl.load_workbook(filepath, read_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    headers = [str(h) if h else "" for h in next(rows_iter)]
    results = []
    for row in rows_iter:
        results.append({h: (v or "") for h, v in zip(headers, row)})
    wb.close()
    return headers, results


# ════════════════════════════════════════════════════════════
# Shopify 转换 + 价格匹配（从 switch_gui.py 搬入）
# ════════════════════════════════════════════════════════════

PRICE_LIBRARY = [
    8.99, 9.01, 9.31, 9.33, 9.37, 9.63, 9.71, 9.72, 9.84, 9.86, 9.95, 9.99,
    10.03, 10.05, 10.11, 10.13, 10.15, 10.35, 10.36, 10.38, 10.54, 10.63, 10.64, 10.66,
    10.71, 10.78, 10.83, 10.94, 10.95, 10.99, 11.22, 11.36, 11.48, 11.65, 11.87, 11.95,
    11.99, 12.22, 12.25, 12.36, 12.41, 12.87, 12.92, 12.95, 12.99, 13.02, 13.13, 13.14,
    13.47, 13.69, 13.95, 13.99, 14.25, 14.32, 14.62, 14.83, 14.91, 14.95, 14.99, 15.31,
    15.52, 15.78, 15.95, 15.99, 16.34, 16.53, 16.66, 16.77, 16.95, 16.99, 17.11, 17.34,
    17.57, 17.95, 17.99, 18.62, 18.81, 18.84, 18.95, 18.99, 19.04, 19.25, 19.95, 19.99,
    20.14, 20.19, 20.34, 20.68, 20.74, 20.95, 20.99, 21.3, 21.47, 21.74, 21.95, 21.99,
    22.31, 22.34, 22.81, 22.95, 22.99, 23.14, 23.32, 23.95, 23.99, 24.22, 24.46, 24.63,
    24.95, 24.99, 25.37, 25.95, 25.99, 26.18, 26.34, 26.95, 26.99, 27.31, 27.54, 27.95,
    27.99, 28.11, 28.95, 28.99, 29, 29.31, 29.95, 29.99, 30.54, 30.95, 30.99, 31.26,
    31.52, 31.95, 31.99, 32.14, 32.24, 32.95, 32.99, 33.62, 33.84, 33.95, 33.99, 34.18,
    34.62, 34.95, 34.99, 35, 35.14, 35.33, 35.95, 35.99, 36.47, 36.95, 36.99, 37.21,
    37.95, 37.99, 38.14, 38.52, 38.95, 38.99, 39, 39.54, 39.95, 39.99, 40.02, 40.15,
    40.95, 40.99, 41.32, 41.95, 41.99, 42.51, 42.95, 42.99, 43.16, 43.95, 43.99, 44.82,
    44.95, 44.99, 45.17, 45.95, 45.99, 46.57, 47.82, 48.36, 49, 49.31, 50.03, 50.13,
    50.81, 55, 59, 69, 79, 85, 89, 99, 109, 119, 129, 189, 219, 259, 299, 319,
    349, 369, 399, 459, 499, 519, 599,
]

SHOPIFY_COLUMNS = [
    "Link-Href", "Handle", "Title", "Body (HTML)", "Collection",
    "Vendor", "Type", "Tags", "Published", "Option1 Name",
    "Option1 Value", "Option2 Name", "Option2 Value", "Option3 Name",
    "Option3 Value", "Variant SKU", "Variant Grams",
    "Variant Inventory Tracker", "Variant Inventory Qty",
    "Variant Inventory Policy", "Variant Fulfillment Service",
    "Variant Price", "Variant Compare At Price",
    "Variant Requires Shipping", "Variant Taxable",
    "Variant Barcode", "Image Src", "Image Position",
    "Image Alt Text", "Gift Card", "SEO Title",
    "SEO Description", "Google Shopping - Google Product Category",
    "Google Shopping - Gender", "Google Shopping - Age Group",
    "Google Shopping - MPN", "Google Shopping - AdWords Grouping",
    "Google Shopping - AdWords Labels", "Google Shopping - Condition",
    "Google Shopping - Custom Product", "Google Shopping - Custom Label 0",
    "Google Shopping - Custom Label 1", "Google Shopping - Custom Label 2",
    "Google Shopping - Custom Label 3", "Google Shopping - Custom Label 4",
    "Variant Image", "Variant Weight Unit", "Variant Tax Code",
    "Cost per item",
]


def _clean_cell(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def _detect_encoding(path):
    for enc in ("utf-8-sig", "gbk", "latin-1"):
        try:
            with open(path, encoding=enc, newline="") as f:
                f.read(2048)
            return enc
        except UnicodeDecodeError:
            continue
    return "utf-8-sig"


def _read_table(path):
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(path)
    return pd.read_csv(path, encoding=_detect_encoding(path))


def _save_table(df, path):
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xls"):
        df.to_excel(path, index=False)
    else:
        df.to_csv(path, index=False, encoding="utf-8-sig")


def _clean_handle_name(name):
    name = _clean_cell(name)
    cleaned = re.sub(r"[^a-zA-Z\s]", "", name).lower()
    return cleaned.replace(" ", "-").strip("-")


def _generate_unique_handle(base, used_handles):
    while True:
        digits = "".join(random.choices(string.digits, k=3))
        letters1 = "".join(random.choices(string.ascii_lowercase, k=2))
        letters2 = "".join(random.choices(string.ascii_lowercase, k=2))
        handle = f"{base}-{letters1}{digits}{letters2}" if base else f"-{letters1}{digits}{letters2}"
        if handle not in used_handles:
            used_handles.add(handle)
            return handle


def _generate_variant_sku(used_skus=None):
    while True:
        first_digit = random.choice("123456789")
        digits1 = "".join(random.choices(string.digits, k=2))
        letters1 = "".join(random.choices(string.ascii_uppercase, k=2))
        letter1 = random.choice(string.ascii_uppercase)
        digits2 = "".join(random.choices(string.digits, k=2))
        letters2 = "".join(random.choices(string.ascii_uppercase, k=2))
        digits3 = "".join(random.choices(string.digits, k=3))
        sku = f"{first_digit}{digits1}{letters1}-{letter1}{digits2}{letters2}-{digits3}{letters1}"
        if used_skus is None or sku not in used_skus:
            if used_skus is not None:
                used_skus.add(sku)
            return sku


def _parse_style_group(value):
    parts = [part.strip() for part in _style_split(_clean_cell(value), "#") if part.strip()]
    if not parts:
        return "", [""]
    if len(parts) == 1:
        return _style_unescape(parts[0]), [""]
    return _style_unescape(parts[0]), parts[1:]


def _parse_variant_segment(segment):
    text = _clean_cell(segment)
    image = ""
    text, image, has_image = _style_split_first(text, "@")
    if has_image:
        image = image.strip()
    price1 = ""
    price2 = ""
    price_parts = text.rsplit("$", 2)
    if len(price_parts) == 3:
        text, price1, price2 = price_parts
    option_parts = _style_split(text, "&") if text else []

    def _field(index):
        return _style_unescape(option_parts[index]) if len(option_parts) > index else ""

    return {
        "option1_value": _field(0),
        "option2_name": _field(1),
        "option2_value": _field(2),
        "option3_name": _field(3),
        "option3_value": _field(4),
        "price1": price1,
        "price2": price2,
        "image": image,
    }


def _styles_to_shopify(input_file, log):
    """将 styles xlsx 转换为 Shopify 格式并自动价格匹配。返回价格匹配文件路径。"""
    input_path = Path(input_file)
    df = pd.read_excel(input_path)
    required = ["title", "name", "price1", "price2", "details", "src_links", "styles1", "styles2", "styles3"]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"缺少必要列: {', '.join(missing)}")

    rows = []
    used_handles = set()
    used_skus = set()

    for _, item in df.iterrows():
        base_handle = _clean_handle_name(item.get("name", ""))
        handle = _generate_unique_handle(base_handle, used_handles)

        option1_name, option1_values = _parse_style_group(item.get("styles1", ""))
        option2_name, option2_values = _parse_style_group(item.get("styles2", ""))
        option3_name, option3_values = _parse_style_group(item.get("styles3", ""))
        if option1_name.startswith("$"):
            option1_name = ""
        if option2_name.startswith("$"):
            option2_name = ""
        if option3_name.startswith("$"):
            option3_name = ""
        images = [img.strip() for img in _clean_cell(item.get("src_links", "")).split("#") if img.strip()]
        combinations = list(itertools_product(option1_values or [""], option2_values or [""], option3_values or [""]))
        total_rows = max(len(combinations), len(images), 1)

        for index in range(total_rows):
            row = {col: "" for col in SHOPIFY_COLUMNS}
            row["Link-Href"] = _clean_cell(item.get("link-href", ""))
            row["Handle"] = handle
            row["Variant Weight Unit"] = "kg"

            if index == 0:
                row["Title"] = _clean_cell(item.get("name", ""))
                row["Body (HTML)"] = _clean_cell(item.get("details", ""))
                row["Collection"] = _clean_cell(item.get("title", ""))
                row["Vendor"] = _clean_cell(item.get("title", ""))
                row["Published"] = "TRUE"
                row["Gift Card"] = "FALSE"
                row["SEO Title"] = row["Title"]
                row["SEO Description"] = row["Title"]

            if index < len(combinations):
                combo = combinations[index]
                segment_info = _parse_variant_segment(combo[0])
                row["Option1 Name"] = option1_name
                row["Option1 Value"] = segment_info["option1_value"] or _clean_cell(combo[0])
                row["Option2 Name"] = segment_info["option2_name"] or option2_name
                row["Option2 Value"] = segment_info["option2_value"] or _clean_cell(combo[1])
                row["Option3 Name"] = segment_info["option3_name"] or option3_name
                row["Option3 Value"] = segment_info["option3_value"] or _clean_cell(combo[2])
                row["Variant SKU"] = _generate_variant_sku(used_skus)
                row["Variant Grams"] = 0
                row["Variant Inventory Tracker"] = "shopify"
                row["Variant Inventory Qty"] = 999
                row["Variant Inventory Policy"] = "deny"
                row["Variant Fulfillment Service"] = "manual"
                row["Variant Price"] = segment_info["price1"] or _clean_cell(item.get("price1", ""))
                row["Variant Compare At Price"] = segment_info["price2"] or _clean_cell(item.get("price2", ""))
                row["Variant Requires Shipping"] = "TRUE"
                row["Variant Taxable"] = "FALSE"
                row["Variant Image"] = segment_info["image"]

            if index < len(images):
                row["Image Src"] = images[index]
                row["Image Position"] = index + 1

            rows.append(row)

    xlsx_dir = input_path.parent / "xlsx"
    csv_dir = input_path.parent / "csv"
    xlsx_dir.mkdir(exist_ok=True)
    csv_dir.mkdir(exist_ok=True)

    output_path = xlsx_dir / f"{input_path.stem}_Shopify_原价.xlsx"
    result_df = pd.DataFrame(rows, columns=SHOPIFY_COLUMNS)
    result_df.to_excel(output_path, index=False)
    csv_path = csv_dir / f"{input_path.stem}_Shopify_原价.csv"
    result_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    log(f"已生成 Shopify 文件: {output_path}")
    log(f"已生成 CSV 文件: {csv_path}")
    log(f"商品数: {len(df)}，输出行数: {len(rows)}")

    log("")
    log("开始价格匹配...")
    price_match_path = _price_match(str(output_path), log)

    # ── Shopify 转 WP ──
    log("")
    log("开始 Shopify 转 WP...")
    try:
        _shopify_to_wp(str(csv_path), log)
        if price_match_path:
            price_csv = Path(str(price_match_path).replace(str(price_match_path).split(".")[-1], "csv"))
            if price_csv.exists():
                _shopify_to_wp(str(price_csv), log)
        log("Shopify 转 WP 完成")
    except Exception as exc:
        log(f"Shopify 转 WP 出错: {exc}")
        import traceback
        traceback.print_exc()

    return output_path


def _price_match(input_file, log):
    """对 Shopify 文件执行价格匹配。"""
    input_path = Path(input_file)
    df = _read_table(input_path).reset_index(drop=True)
    for col in ("Handle", "Variant Price"):
        if col not in df.columns:
            raise ValueError(f"缺少必要列: {col}")

    source_stem = input_path.stem
    for suffix in ("_Shopify_待匹配", "_Shopify_原价_拆分", "_Shopify_原价", "_原价"):
        if source_stem.endswith(suffix):
            source_stem = source_stem[: -len(suffix)]
            break

    # 使用 object 类型，既保留原始空值/文本，也允许写入小数价格。
    df["Variant Price"] = df["Variant Price"].astype(object)
    if "Variant Compare At Price" not in df.columns:
        df["Variant Compare At Price"] = None
    df["Variant Compare At Price"] = df["Variant Compare At Price"].astype(object)
    df["匹配价格"] = pd.Series([None] * len(df), dtype=object)

    global_used_prices = set()
    price_library_sorted = sorted(PRICE_LIBRARY)

    for handle, group in df.groupby("Handle"):
        if pd.isna(handle):
            continue
        valid_rows = []
        for idx, row in group.iterrows():
            price = pd.to_numeric(row.get("Variant Price"), errors="coerce")
            if pd.notna(price):
                valid_rows.append((idx, float(price)))
        if not valid_rows:
            continue

        available_prices = [p for p in price_library_sorted if p not in global_used_prices]
        for idx, original_price in sorted(valid_rows, key=lambda x: x[1]):
            lower_bound = original_price * 0.75
            upper_bound = original_price * 1.25
            candidates = [p for p in available_prices if lower_bound <= p <= upper_bound]
            if candidates:
                matched = candidates[0]
                compare_at_price = round(matched * 1.2, 2)

                global_used_prices.add(matched)
                available_prices.remove(matched)

                # 保留匹配结果，方便核对价格匹配情况
                df.at[idx, "匹配价格"] = matched

                # 匹配成功后，直接更新 Shopify 售价与划线原价
                df.at[idx, "Variant Price"] = matched
                df.at[idx, "Variant Compare At Price"] = compare_at_price

    if input_path.parent.name == "xlsx":
        base_dir = input_path.parent.parent
    else:
        base_dir = input_path.parent
    xlsx_dir = base_dir / "xlsx"
    csv_dir = base_dir / "csv"
    xlsx_dir.mkdir(exist_ok=True)
    csv_dir.mkdir(exist_ok=True)

    output_path = xlsx_dir / f"{source_stem}_Shopify_价格匹配{input_path.suffix}"
    _save_table(df, output_path)
    csv_output = csv_dir / f"{source_stem}_Shopify_价格匹配.csv"
    df.to_csv(csv_output, index=False, encoding="utf-8-sig")
    log(f"已生成价格匹配 CSV: {csv_output}")

    unused_prices = [p for p in price_library_sorted if p not in global_used_prices]
    if unused_prices:
        unused_df = pd.DataFrame({"未匹配价格": unused_prices})
        unused_xlsx = xlsx_dir / f"{source_stem}_Shopify_未匹配到的价格{input_path.suffix}"
        _save_table(unused_df, unused_xlsx)
        unused_csv = csv_dir / f"{source_stem}_Shopify_未匹配到的价格.csv"
        unused_df.to_csv(unused_csv, index=False, encoding="utf-8-sig")
        log(f"未匹配价格文件: {unused_xlsx}")

    log(f"已生成价格匹配文件: {output_path}")
    log(f"已使用价格数: {len(global_used_prices)}，未使用价格数: {len(unused_prices)}")
    return output_path


# ════════════════════════════════════════════════════════════
# Shopify 转 WP（从 switch_gui.py 搬入）
# ════════════════════════════════════════════════════════════

WP_HEADERS = [
    "Type", "SKU", "Name", "Published", "Visibility in catalog", "Description", "In stock?", "Stock",
    "Sale price", "Regular price", "Categories", "Tags", "Images", "Parent", "Position",
    "Attribute 1 name", "Attribute 1 value(s)", "Attribute 1 visible", "Attribute 1 global",
    "Attribute 2 name", "Attribute 2 value(s)", "Attribute 2 visible", "Attribute 2 global",
    "Attribute 3 name", "Attribute 3 value(s)", "Attribute 3 visible", "Attribute 3 global",
]


def _generate_parent_sku(handle="", counter=0, used_skus=None):
    """基于商品 Handle 生成确定性 parent SKU，同一 Handle 始终返回相同 SKU。"""
    while True:
        hash_input = f"{handle}_{counter}" if handle else str(counter)
        h = hashlib.md5(hash_input.encode()).hexdigest()
        digits = "".join(c for c in h if c.isdigit())
        if len(digits) < 8:
            digits = (digits + "0" * 8)[:8]
        sku = f"SKU{digits[:8]}"
        if used_skus is None or sku not in used_skus:
            if used_skus is not None:
                used_skus.add(sku)
            return sku
        counter += 1


def _wp_title(row):
    values = [row.get("Option1 Value", ""), row.get("Option2 Value", ""), row.get("Option3 Value", "")]
    values = [value for value in values if value]
    return "-" + ",".join(values) if values else ""


def _csv_join_unique(existing, value):
    value = _clean_cell(value)
    items = [item.strip() for item in _clean_cell(existing).split(",") if item.strip()]
    if value and value not in items:
        items.append(value)
    return ",".join(items)


def _read_csv_rows(path):
    encoding = _detect_encoding(path)
    with open(path, encoding=encoding, newline="") as f:
        return list(csv.DictReader(f)), encoding


def _read_shopify_rows(path):
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xls"):
        df = pd.read_excel(path).fillna("")
        return [{key: _clean_cell(value) for key, value in row.items()} for row in df.to_dict("records")]
    rows, _ = _read_csv_rows(path)
    return rows


def _shopify_to_wp(input_file, log):
    """将 Shopify CSV 转换为 WooCommerce (WP) 格式 CSV。"""
    input_path = Path(input_file)
    rows = _read_shopify_rows(input_path)
    products = {}
    used_parent_skus = set()

    for row in rows:
        handle = row.get("Handle", "").strip()
        if not handle:
            continue
        if handle not in products:
            parent = {}
            is_variable = row.get("Option1 Value", "") not in ("", "Default Title")
            parent["Type"] = "variable" if is_variable else "simple"
            parent["SKU"] = (
                _generate_parent_sku(handle, used_skus=used_parent_skus)
                if is_variable
                else (row.get("Variant SKU", "") or _generate_parent_sku(handle, used_skus=used_parent_skus))
            )
            parent["Name"] = row.get("Title", "")
            parent["Published"] = "1"
            parent["Visibility in catalog"] = "visible"
            parent["Description"] = row.get("Body (HTML)", "")
            parent["In stock?"] = "1"
            parent["Stock"] = "99999"
            parent["Sale price"] = "" if is_variable else row.get("Variant Price", "")
            parent["Regular price"] = "" if is_variable else (row.get("Variant Compare At Price", "") or row.get("Variant Price", ""))
            parent["Categories"] = row.get("Collection", "") or row.get("Type", "")
            parent["Tags"] = row.get("Tags", "")
            parent["Images"] = row.get("Image Src", "")
            parent["Parent"] = ""
            parent["Position"] = "0"
            raw_opt1_name = row.get("Option1 Name", "")
            opt1_name = "" if (raw_opt1_name == "Title" or raw_opt1_name.startswith("$")) else raw_opt1_name
            parent["Attribute 1 name"] = opt1_name
            parent["Attribute 1 value(s)"] = "" if row.get("Option1 Value", "") == "Default Title" else row.get("Option1 Value", "")
            parent["Attribute 1 visible"] = "1" if parent["Attribute 1 name"] else ""
            parent["Attribute 1 global"] = "1" if parent["Attribute 1 name"] else ""
            parent["Attribute 2 name"] = row.get("Option2 Name", "")
            parent["Attribute 2 value(s)"] = row.get("Option2 Value", "")
            parent["Attribute 2 visible"] = "1" if row.get("Option2 Name", "") else ""
            parent["Attribute 2 global"] = "1" if row.get("Option2 Name", "") else ""
            parent["Attribute 3 name"] = row.get("Option3 Name", "")
            parent["Attribute 3 value(s)"] = row.get("Option3 Value", "")
            parent["Attribute 3 visible"] = "1" if row.get("Option3 Name", "") else ""
            parent["Attribute 3 global"] = "1" if row.get("Option3 Name", "") else ""
            products[handle] = [parent]

        parent = products[handle][0]
        if row.get("Image Src", ""):
            parent["Images"] = _csv_join_unique(parent["Images"], row.get("Image Src", ""))

        if row.get("Option1 Value", "") and row.get("Option1 Value", "") != "Default Title":
            parent["Attribute 1 value(s)"] = _csv_join_unique(
                parent["Attribute 1 value(s)"], row.get("Option1 Value", "")
            )
            parent["Attribute 2 value(s)"] = _csv_join_unique(
                parent["Attribute 2 value(s)"], row.get("Option2 Value", "")
            )
            parent["Attribute 3 value(s)"] = _csv_join_unique(
                parent["Attribute 3 value(s)"], row.get("Option3 Value", "")
            )
            variation = {
                "Type": "variation",
                "SKU": row.get("Variant SKU", ""),
                "Name": parent["Name"] + _wp_title(row),
                "Published": "1",
                "Visibility in catalog": "visible",
                "Description": "",
                "In stock?": "1",
                "Stock": "99999",
                "Sale price": row.get("Variant Price", ""),
                "Regular price": row.get("Variant Compare At Price", "") or row.get("Variant Price", ""),
                "Categories": parent.get("Categories", ""),
                "Tags": "",
                "Images": row.get("Variant Image", ""),
                "Parent": parent["SKU"],
                "Position": str(len(products[handle])),
                "Attribute 1 name": parent["Attribute 1 name"],
                "Attribute 1 value(s)": row.get("Option1 Value", ""),
                "Attribute 1 visible": "",
                "Attribute 1 global": "1" if row.get("Option1 Value", "") else "",
                "Attribute 2 name": parent["Attribute 2 name"],
                "Attribute 2 value(s)": row.get("Option2 Value", ""),
                "Attribute 2 visible": "",
                "Attribute 2 global": "1" if row.get("Option2 Value", "") else "",
                "Attribute 3 name": parent["Attribute 3 name"],
                "Attribute 3 value(s)": row.get("Option3 Value", ""),
                "Attribute 3 visible": "",
                "Attribute 3 global": "1" if row.get("Option3 Value", "") else "",
            }
            products[handle].append(variation)

    # ── 价格校验：Sale price 不能大于 Regular price ──
    price_fix_count = 0
    for items in products.values():
        for item in items:
            try:
                sale = float(item.get("Sale price", "") or 0)
                regular = float(item.get("Regular price", "") or 0)
                if sale > 0 and regular > 0 and sale > regular:
                    item["Regular price"] = item["Sale price"]
                    price_fix_count += 1
            except (ValueError, TypeError):
                pass
    if price_fix_count:
        log(f"价格修正: {price_fix_count} 行 Sale price > Regular price，已统一为 Sale price")

    base_stem = input_path.stem
    for suffix in ("_Shopify_价格匹配", "_Shopify_原价"):
        if base_stem.endswith(suffix):
            base_stem = base_stem[: -len(suffix)]
            break
    output_dir = input_path.parent / f"wp-{base_stem}"
    output_dir.mkdir(exist_ok=True)
    output_path = output_dir / f"wp-{input_path.stem}.csv"
    wp_headers_no_tags = [h for h in WP_HEADERS if h != "Tags"]
    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=wp_headers_no_tags)
        writer.writeheader()
        for items in products.values():
            for item in items:
                writer.writerow({header: item.get(header, "") for header in wp_headers_no_tags})

    total_rows = sum(len(items) for items in products.values())
    log(f"已生成 WP 文件: {output_path}")
    log(f"商品组数: {len(products)}，输出行数: {total_rows}")
    return output_path


# ════════════════════════════════════════════════════════════
# 图形界面（清爽蓝白后台风）
# ════════════════════════════════════════════════════════════

UI_COLORS = {
    "background": "#F4F7FB",
    "card": "#FFFFFF",
    "card_alt": "#F8FAFC",
    "border": "#E2E8F0",
    "border_focus": "#93C5FD",
    "text": "#0F172A",
    "text_secondary": "#475569",
    "text_muted": "#94A3B8",
    "primary": "#2563EB",
    "primary_hover": "#1D4ED8",
    "primary_light": "#EFF6FF",
    "primary_border": "#BFDBFE",
    "success": "#16A34A",
    "success_light": "#F0FDF4",
    "success_border": "#BBF7D0",
    "warning": "#D97706",
    "warning_light": "#FFFBEB",
    "warning_border": "#FDE68A",
    "danger": "#DC2626",
    "danger_hover": "#B91C1C",
    "danger_light": "#FEF2F2",
    "danger_border": "#FECACA",
    "disabled": "#CBD5E1",
    "log_background": "#0F172A",
    "log_panel": "#111827",
    "log_text": "#D1D5DB",
}

UI_FONT = "Microsoft YaHei UI"
MONO_FONT = "Consolas"


class TextRedirector:
    """将 print 输出安全地转发到日志文本框。"""

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


class ScraperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Shopify 商品采集工作台")
        self.root.geometry("980x760")
        self.root.minsize(900, 680)
        self.root.configure(bg=UI_COLORS["background"])

        self.style = ttk.Style()
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")
        self._configure_styles()

        self.is_running = False
        self.stop_event = threading.Event()
        self.pending_merge = None
        self.file_rows = []
        self.task_files = []
        self.task_folder = ""
        self.active_file_index = None

        self.folder_var = tk.StringVar()
        self.status_var = tk.StringVar(value="请选择分类文件夹后点击“读取文件夹”")
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_percent_var = tk.StringVar(value="0%")
        self.file_count_var = tk.StringVar(value="0 个文件")

        self._build_log_window()
        sys.stdout = TextRedirector(self.log_text)
        sys.stderr = sys.stdout
        self._build_ui()
        self._center_window()

    # ────────────────────────────────────────────────────────
    # 样式与通用组件
    # ────────────────────────────────────────────────────────
    def _configure_styles(self):
        self.style.configure(
            "App.TEntry",
            fieldbackground=UI_COLORS["card_alt"],
            foreground=UI_COLORS["text"],
            bordercolor=UI_COLORS["border"],
            lightcolor=UI_COLORS["border"],
            darkcolor=UI_COLORS["border"],
            insertcolor=UI_COLORS["primary"],
            padding=(10, 9),
            font=(UI_FONT, 10),
        )
        self.style.map(
            "App.TEntry",
            bordercolor=[("focus", UI_COLORS["border_focus"])],
            lightcolor=[("focus", UI_COLORS["border_focus"])],
            darkcolor=[("focus", UI_COLORS["border_focus"])],
        )
        self.style.configure(
            "Blue.Horizontal.TProgressbar",
            troughcolor="#E8EEF7",
            background=UI_COLORS["primary"],
            bordercolor="#E8EEF7",
            lightcolor=UI_COLORS["primary"],
            darkcolor=UI_COLORS["primary"],
            thickness=10,
        )
        self.style.configure(
            "App.Vertical.TScrollbar",
            troughcolor=UI_COLORS["card"],
            background="#CBD5E1",
            bordercolor=UI_COLORS["card"],
            arrowcolor=UI_COLORS["text_secondary"],
            gripcount=0,
        )
        self.style.map(
            "App.Vertical.TScrollbar",
            background=[("active", "#94A3B8")],
        )

    def _center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = max((screen_width - width) // 2, 0)
        y = max((screen_height - height) // 2 - 20, 0)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _create_card(self, parent, **pack_kwargs):
        card = tk.Frame(
            parent,
            bg=UI_COLORS["card"],
            highlightbackground=UI_COLORS["border"],
            highlightcolor=UI_COLORS["border"],
            highlightthickness=1,
            bd=0,
        )
        card.pack(**pack_kwargs)
        return card

    def _create_button(
        self,
        parent,
        text,
        command,
        kind="secondary",
        width=None,
        padx=16,
        pady=9,
        font_size=10,
    ):
        styles = {
            "primary": {
                "bg": UI_COLORS["primary"],
                "fg": "#FFFFFF",
                "hover": UI_COLORS["primary_hover"],
                "disabled_bg": "#AFC7F5",
                "disabled_fg": "#F8FAFC",
                "border": 0,
            },
            "danger": {
                "bg": UI_COLORS["card"],
                "fg": UI_COLORS["danger"],
                "hover": UI_COLORS["danger_light"],
                "disabled_bg": UI_COLORS["card_alt"],
                "disabled_fg": UI_COLORS["text_muted"],
                "border": 1,
            },
            "secondary": {
                "bg": UI_COLORS["card_alt"],
                "fg": UI_COLORS["text_secondary"],
                "hover": "#EEF2F7",
                "disabled_bg": "#F1F5F9",
                "disabled_fg": UI_COLORS["text_muted"],
                "border": 1,
            },
            "ghost": {
                "bg": UI_COLORS["card"],
                "fg": UI_COLORS["primary"],
                "hover": UI_COLORS["primary_light"],
                "disabled_bg": UI_COLORS["card"],
                "disabled_fg": UI_COLORS["text_muted"],
                "border": 1,
            },
        }
        palette = styles[kind]
        button = tk.Button(
            parent,
            text=text,
            command=command,
            width=width,
            padx=padx,
            pady=pady,
            font=(UI_FONT, font_size, "bold" if kind == "primary" else "normal"),
            bg=palette["bg"],
            fg=palette["fg"],
            activebackground=palette["hover"],
            activeforeground=palette["fg"],
            disabledforeground=palette["disabled_fg"],
            relief=tk.FLAT,
            bd=0,
            highlightthickness=palette["border"],
            highlightbackground=(
                UI_COLORS["danger_border"]
                if kind == "danger"
                else UI_COLORS["border"]
            ),
            highlightcolor=(
                UI_COLORS["danger_border"]
                if kind == "danger"
                else UI_COLORS["border"]
            ),
            cursor="hand2",
        )
        button._ui_kind = kind
        button._normal_bg = palette["bg"]
        button._hover_bg = palette["hover"]
        button._disabled_bg = palette["disabled_bg"]
        button.bind("<Enter>", lambda _e, b=button: self._button_hover(b, True))
        button.bind("<Leave>", lambda _e, b=button: self._button_hover(b, False))
        return button

    @staticmethod
    def _button_hover(button, entering):
        if str(button.cget("state")) == str(tk.DISABLED):
            return
        button.configure(bg=button._hover_bg if entering else button._normal_bg)

    @staticmethod
    def _set_button_enabled(button, enabled):
        if enabled:
            button.configure(state=tk.NORMAL, bg=button._normal_bg, cursor="hand2")
        else:
            button.configure(state=tk.DISABLED, bg=button._disabled_bg, cursor="arrow")

    # ────────────────────────────────────────────────────────
    # 日志窗口
    # ────────────────────────────────────────────────────────
    def _build_log_window(self):
        self.log_win = tk.Toplevel(self.root)
        self.log_win.title("运行日志")
        self.log_win.geometry("820x480")
        self.log_win.minsize(650, 360)
        self.log_win.configure(bg=UI_COLORS["log_background"])
        self.log_win.protocol("WM_DELETE_WINDOW", self.log_win.withdraw)

        header = tk.Frame(self.log_win, bg=UI_COLORS["log_panel"], height=58)
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        title_group = tk.Frame(header, bg=UI_COLORS["log_panel"])
        title_group.pack(side=tk.LEFT, padx=18, pady=10)
        tk.Label(
            title_group,
            text="运行日志",
            bg=UI_COLORS["log_panel"],
            fg="#F8FAFC",
            font=(UI_FONT, 12, "bold"),
        ).pack(anchor=tk.W)
        tk.Label(
            title_group,
            text="实时查看采集、保存与转换过程",
            bg=UI_COLORS["log_panel"],
            fg="#94A3B8",
            font=(UI_FONT, 9),
        ).pack(anchor=tk.W)

        log_actions = tk.Frame(header, bg=UI_COLORS["log_panel"])
        log_actions.pack(side=tk.RIGHT, padx=14)
        clear_btn = tk.Button(
            log_actions,
            text="清空",
            command=self._clear_logs,
            bg="#1E293B",
            fg="#CBD5E1",
            activebackground="#334155",
            activeforeground="#FFFFFF",
            relief=tk.FLAT,
            bd=0,
            padx=14,
            pady=6,
            font=(UI_FONT, 9),
            cursor="hand2",
        )
        clear_btn.pack(side=tk.LEFT, padx=(0, 8))
        close_btn = tk.Button(
            log_actions,
            text="关闭",
            command=self.log_win.withdraw,
            bg="#1E293B",
            fg="#CBD5E1",
            activebackground="#334155",
            activeforeground="#FFFFFF",
            relief=tk.FLAT,
            bd=0,
            padx=14,
            pady=6,
            font=(UI_FONT, 9),
            cursor="hand2",
        )
        close_btn.pack(side=tk.LEFT)

        log_body = tk.Frame(self.log_win, bg=UI_COLORS["log_background"])
        log_body.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        self.log_text = scrolledtext.ScrolledText(
            log_body,
            state=tk.DISABLED,
            bg=UI_COLORS["log_background"],
            fg=UI_COLORS["log_text"],
            insertbackground="#FFFFFF",
            selectbackground="#1D4ED8",
            selectforeground="#FFFFFF",
            font=(MONO_FONT, 9),
            relief=tk.FLAT,
            bd=0,
            padx=12,
            pady=12,
            wrap=tk.WORD,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_win.withdraw()

    def _clear_logs(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)

    def show_logs(self):
        self.log_win.deiconify()
        self.log_win.lift()
        self.log_win.focus_force()

    # ────────────────────────────────────────────────────────
    # 主界面
    # ────────────────────────────────────────────────────────
    def _build_ui(self):
        page = tk.Frame(self.root, bg=UI_COLORS["background"])
        page.pack(fill=tk.BOTH, expand=True)

        content = tk.Frame(page, bg=UI_COLORS["background"])
        content.pack(fill=tk.BOTH, expand=True, padx=24, pady=(16, 14))

        # 先为底部区域预留空间，避免窗口高度较小时操作按钮被上方内容挤出。
        footer = tk.Frame(content, bg=UI_COLORS["background"])
        footer.pack(side=tk.BOTTOM, fill=tk.X)
        self._build_progress_card(footer)
        self._build_action_bar(footer)

        # 上方主体区域可以随窗口高度变化；文件列表保持紧凑并独立滚动。
        body = tk.Frame(content, bg=UI_COLORS["background"])
        body.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self._build_header(body)
        self._build_folder_card(body)
        self._build_file_card(body)

    def _build_header(self, parent):
        header = tk.Frame(parent, bg=UI_COLORS["background"])
        header.pack(fill=tk.X, pady=(0, 12))

        brand = tk.Frame(header, bg=UI_COLORS["background"])
        brand.pack(side=tk.LEFT)

        logo = tk.Label(
            brand,
            text="S",
            width=3,
            height=1,
            bg=UI_COLORS["primary"],
            fg="#FFFFFF",
            font=(UI_FONT, 15, "bold"),
            relief=tk.FLAT,
        )
        logo.pack(side=tk.LEFT, padx=(0, 12), ipady=6)

        title_group = tk.Frame(brand, bg=UI_COLORS["background"])
        title_group.pack(side=tk.LEFT)
        tk.Label(
            title_group,
            text="Shopify 商品采集工作台",
            bg=UI_COLORS["background"],
            fg=UI_COLORS["text"],
            font=(UI_FONT, 18, "bold"),
        ).pack(anchor=tk.W)
        tk.Label(
            title_group,
            text="批量读取商品链接，自动采集、合并并转换 Shopify 数据",
            bg=UI_COLORS["background"],
            fg=UI_COLORS["text_secondary"],
            font=(UI_FONT, 9),
        ).pack(anchor=tk.W, pady=(3, 0))

        self.header_log_btn = self._create_button(
            header,
            "查看运行日志",
            self.show_logs,
            kind="ghost",
            padx=14,
            pady=8,
            font_size=9,
        )
        self.header_log_btn.pack(side=tk.RIGHT, pady=4)

    def _build_folder_card(self, parent):
        card = self._create_card(parent, fill=tk.X, pady=(0, 12))
        inner = tk.Frame(card, bg=UI_COLORS["card"])
        inner.pack(fill=tk.X, padx=18, pady=14)

        tk.Label(
            inner,
            text="选择分类文件夹",
            bg=UI_COLORS["card"],
            fg=UI_COLORS["text"],
            font=(UI_FONT, 11, "bold"),
        ).pack(anchor=tk.W)
        tk.Label(
            inner,
            text="文件夹中需要包含待处理的 .xlsx 文件",
            bg=UI_COLORS["card"],
            fg=UI_COLORS["text_muted"],
            font=(UI_FONT, 9),
        ).pack(anchor=tk.W, pady=(3, 10))

        input_row = tk.Frame(inner, bg=UI_COLORS["card"])
        input_row.pack(fill=tk.X)

        self.folder_entry = ttk.Entry(
            input_row,
            textvariable=self.folder_var,
            style="App.TEntry",
        )
        self.folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        self.browse_btn = self._create_button(
            input_row,
            "浏览文件夹",
            self._browse_folder,
            kind="secondary",
            padx=14,
            pady=8,
            font_size=9,
        )
        self.browse_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.scan_btn = self._create_button(
            input_row,
            "读取文件夹",
            self._scan_files,
            kind="primary",
            padx=14,
            pady=8,
            font_size=9,
        )
        self.scan_btn.pack(side=tk.LEFT)

    def _build_file_card(self, parent):
        card = self._create_card(
            parent,
            fill=tk.BOTH,
            expand=True,
            pady=(0, 10),
        )
        card.configure(height=250)
        card.pack_propagate(False)
        self.file_card = card

        top = tk.Frame(card, bg=UI_COLORS["card"])
        top.pack(fill=tk.X, padx=18, pady=(10, 6))

        title_group = tk.Frame(top, bg=UI_COLORS["card"])
        title_group.pack(side=tk.LEFT)
        tk.Label(
            title_group,
            text="待处理文件",
            bg=UI_COLORS["card"],
            fg=UI_COLORS["text"],
            font=(UI_FONT, 11, "bold"),
        ).pack(anchor=tk.W)
        tk.Label(
            title_group,
            text="可为每个文件单独设置需要跳过的图片位置",
            bg=UI_COLORS["card"],
            fg=UI_COLORS["text_muted"],
            font=(UI_FONT, 9),
        ).pack(anchor=tk.W, pady=(3, 0))

        count_chip = tk.Label(
            top,
            textvariable=self.file_count_var,
            bg=UI_COLORS["primary_light"],
            fg=UI_COLORS["primary"],
            font=(UI_FONT, 9, "bold"),
            padx=10,
            pady=5,
        )
        count_chip.pack(side=tk.RIGHT)

        table_header = tk.Frame(card, bg=UI_COLORS["card_alt"], height=30)
        table_header.pack(fill=tk.X, padx=18)
        table_header.pack_propagate(False)
        table_header.grid_columnconfigure(1, weight=1)

        headers = [
            ("序号", 0, 54, tk.CENTER),
            ("文件名", 1, 0, tk.W),
            ("保留图片", 2, 100, tk.W),
            ("跳过图片", 3, 100, tk.W),
            ("状态", 4, 100, tk.CENTER),
        ]
        for text, column, width, anchor in headers:
            label = tk.Label(
                table_header,
                text=text,
                bg=UI_COLORS["card_alt"],
                fg=UI_COLORS["text_secondary"],
                font=(UI_FONT, 9, "bold"),
                anchor=anchor,
            )
            label.grid(
                row=0,
                column=column,
                sticky="nsew" if column == 1 else "ns",
                padx=(12 if column == 1 else 6),
            )
            if width:
                label.configure(width=max(width // 9, 1))

        list_shell = tk.Frame(card, bg=UI_COLORS["card"])
        list_shell.pack(fill=tk.BOTH, expand=True, padx=18, pady=(0, 8))

        self._canvas = tk.Canvas(
            list_shell,
            bg=UI_COLORS["card"],
            highlightthickness=0,
            bd=0,
        )
        scrollbar = ttk.Scrollbar(
            list_shell,
            orient=tk.VERTICAL,
            command=self._canvas.yview,
            style="App.Vertical.TScrollbar",
        )
        self.list_inner = tk.Frame(self._canvas, bg=UI_COLORS["card"])
        self._list_window = self._canvas.create_window(
            (0, 0),
            window=self.list_inner,
            anchor="nw",
        )

        self.list_inner.bind(
            "<Configure>",
            lambda _e: self._canvas.configure(scrollregion=self._canvas.bbox("all")),
        )
        self._canvas.bind(
            "<Configure>",
            lambda e: self._canvas.itemconfigure(self._list_window, width=e.width),
        )
        self._canvas.configure(yscrollcommand=scrollbar.set)

        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self._canvas.bind("<Enter>", self._bind_mousewheel)
        self._canvas.bind("<Leave>", self._unbind_mousewheel)
        self.list_inner.bind("<Enter>", self._bind_mousewheel)
        self.list_inner.bind("<Leave>", self._unbind_mousewheel)

        self._show_empty_file_state()

    def _build_progress_card(self, parent):
        card = self._create_card(parent, fill=tk.X, pady=(0, 12))
        inner = tk.Frame(card, bg=UI_COLORS["card"])
        inner.pack(fill=tk.X, padx=18, pady=14)

        progress_header = tk.Frame(inner, bg=UI_COLORS["card"])
        progress_header.pack(fill=tk.X)
        tk.Label(
            progress_header,
            text="任务进度",
            bg=UI_COLORS["card"],
            fg=UI_COLORS["text"],
            font=(UI_FONT, 10, "bold"),
        ).pack(side=tk.LEFT)
        tk.Label(
            progress_header,
            textvariable=self.progress_percent_var,
            bg=UI_COLORS["card"],
            fg=UI_COLORS["primary"],
            font=(UI_FONT, 11, "bold"),
        ).pack(side=tk.RIGHT)

        self.progress_bar = ttk.Progressbar(
            inner,
            orient=tk.HORIZONTAL,
            mode="determinate",
            variable=self.progress_var,
            style="Blue.Horizontal.TProgressbar",
        )
        self.progress_bar.pack(fill=tk.X, pady=(10, 8), ipady=1)

        status_row = tk.Frame(inner, bg=UI_COLORS["card"])
        status_row.pack(fill=tk.X)
        self.status_dot = tk.Label(
            status_row,
            text="●",
            bg=UI_COLORS["card"],
            fg=UI_COLORS["text_muted"],
            font=(UI_FONT, 8),
        )
        self.status_dot.pack(side=tk.LEFT, padx=(0, 7))
        self.status_label = tk.Label(
            status_row,
            textvariable=self.status_var,
            bg=UI_COLORS["card"],
            fg=UI_COLORS["text_secondary"],
            font=(UI_FONT, 9),
            anchor=tk.W,
        )
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

    def _build_action_bar(self, parent):
        actions = tk.Frame(parent, bg=UI_COLORS["background"])
        actions.pack(fill=tk.X)

        hint = tk.Label(
            actions,
            text="开始后可随时打开日志窗口查看详细信息",
            bg=UI_COLORS["background"],
            fg=UI_COLORS["text_muted"],
            font=(UI_FONT, 9),
        )
        hint.pack(side=tk.LEFT)

        button_group = tk.Frame(actions, bg=UI_COLORS["background"])
        button_group.pack(side=tk.RIGHT)

        self.log_btn = self._create_button(
            button_group,
            "查看日志",
            self.show_logs,
            kind="secondary",
            padx=16,
            pady=9,
        )
        self.log_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.stop_btn = self._create_button(
            button_group,
            "停止任务",
            self._stop,
            kind="danger",
            padx=16,
            pady=9,
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._set_button_enabled(self.stop_btn, False)

        self.start_btn = self._create_button(
            button_group,
            "开始采集",
            self._start,
            kind="primary",
            padx=22,
            pady=9,
        )
        self.start_btn.pack(side=tk.LEFT)

        self.merge_btn = self._create_button(
            button_group,
            "合并并转换",
            self._do_merge,
            kind="primary",
            padx=22,
            pady=9,
        )
        self.merge_btn.pack(side=tk.LEFT, padx=(8, 0))
        self._set_button_enabled(self.merge_btn, False)

    # ────────────────────────────────────────────────────────
    # 文件列表交互
    # ────────────────────────────────────────────────────────
    def _bind_mousewheel(self, _event=None):
        self.root.bind_all("<MouseWheel>", self._on_mousewheel)
        self.root.bind_all("<Button-4>", self._on_mousewheel_linux)
        self.root.bind_all("<Button-5>", self._on_mousewheel_linux)

    def _unbind_mousewheel(self, _event=None):
        self.root.unbind_all("<MouseWheel>")
        self.root.unbind_all("<Button-4>")
        self.root.unbind_all("<Button-5>")

    def _on_mousewheel(self, event):
        if event.delta:
            self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_mousewheel_linux(self, event):
        direction = -1 if event.num == 4 else 1
        self._canvas.yview_scroll(direction, "units")

    def _show_empty_file_state(self):
        for widget in self.list_inner.winfo_children():
            widget.destroy()
        empty = tk.Frame(self.list_inner, bg=UI_COLORS["card"])
        empty.pack(fill=tk.BOTH, expand=True, pady=18)
        tk.Label(
            empty,
            text="尚未读取文件",
            bg=UI_COLORS["card"],
            fg=UI_COLORS["text_secondary"],
            font=(UI_FONT, 10, "bold"),
        ).pack()
        tk.Label(
            empty,
            text="选择分类文件夹后，文件会显示在这里",
            bg=UI_COLORS["card"],
            fg=UI_COLORS["text_muted"],
            font=(UI_FONT, 9),
        ).pack(pady=(5, 0))

    def _browse_folder(self):
        path = filedialog.askdirectory(title="选择分类文件夹（如 lindvs.com连衣裙）")
        if path:
            self.folder_var.set(path)
            self._scan_files()

    def _scan_files(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("路径无效", "请先选择一个有效的分类文件夹。")
            return

        xlsx_files = sorted(
            f
            for f in os.listdir(folder)
            if f.lower().endswith(".xlsx") and not f.startswith("~")
        )

        if not xlsx_files:
            self.file_rows.clear()
            self.file_count_var.set("0 个文件")
            self._show_empty_file_state()
            self.update_status(0, 100, "该文件夹中没有找到 .xlsx 文件")
            messagebox.showinfo("未找到文件", "该文件夹下没有 .xlsx 文件。")
            return

        for widget in self.list_inner.winfo_children():
            widget.destroy()
        self.file_rows.clear()

        self.list_inner.grid_columnconfigure(0, minsize=54)
        self.list_inner.grid_columnconfigure(1, weight=1)
        self.list_inner.grid_columnconfigure(2, minsize=100)
        self.list_inner.grid_columnconfigure(3, minsize=100)
        self.list_inner.grid_columnconfigure(4, minsize=100)

        for index, filename in enumerate(xlsx_files):
            file_path = os.path.join(folder, filename)
            keep_var = tk.StringVar()
            skip_var = tk.StringVar()
            status_var = tk.StringVar(value="等待中")

            row_bg = UI_COLORS["card"] if index % 2 == 0 else "#FBFCFE"
            row = tk.Frame(
                self.list_inner,
                bg=row_bg,
                height=48,
                highlightbackground="#EEF2F7",
                highlightthickness=0,
            )
            row.grid(row=index, column=0, columnspan=5, sticky="ew")
            row.grid_columnconfigure(1, weight=1)
            row.grid_propagate(False)

            number_label = tk.Label(
                row,
                text=f"{index + 1:02d}",
                bg=row_bg,
                fg=UI_COLORS["text_muted"],
                font=(MONO_FONT, 9),
                anchor=tk.CENTER,
            )
            number_label.grid(row=0, column=0, sticky="nsew", padx=6)

            file_label = tk.Label(
                row,
                text=filename,
                bg=row_bg,
                fg=UI_COLORS["text"],
                font=(UI_FONT, 9),
                anchor=tk.W,
            )
            file_label.grid(row=0, column=1, sticky="nsew", padx=(12, 12))

            # 保留图片
            keep_shell = tk.Frame(row, bg=row_bg, width=100)
            keep_shell.grid(row=0, column=2, sticky="w", padx=(0, 6))
            keep_entry = tk.Entry(
                keep_shell,
                textvariable=keep_var,
                width=7,
                bg="#FFFFFF",
                fg=UI_COLORS["text"],
                insertbackground=UI_COLORS["primary"],
                relief=tk.FLAT,
                bd=0,
                highlightthickness=1,
                highlightbackground=UI_COLORS["primary_border"],
                highlightcolor=UI_COLORS["border_focus"],
                font=(UI_FONT, 9),
            )
            keep_entry.pack(side=tk.LEFT, ipady=5)
            tk.Label(
                keep_shell,
                text="如 1,2",
                bg=row_bg,
                fg=UI_COLORS["text_muted"],
                font=(UI_FONT, 8),
            ).pack(side=tk.LEFT, padx=(4, 0))

            # 跳过图片
            skip_shell = tk.Frame(row, bg=row_bg, width=100)
            skip_shell.grid(row=0, column=3, sticky="w", padx=(0, 6))
            skip_entry = tk.Entry(
                skip_shell,
                textvariable=skip_var,
                width=7,
                bg="#FFFFFF",
                fg=UI_COLORS["text"],
                insertbackground=UI_COLORS["primary"],
                relief=tk.FLAT,
                bd=0,
                highlightthickness=1,
                highlightbackground=UI_COLORS["border"],
                highlightcolor=UI_COLORS["border_focus"],
                font=(UI_FONT, 9),
            )
            skip_entry.pack(side=tk.LEFT, ipady=5)
            tk.Label(
                skip_shell,
                text="如 1,3",
                bg=row_bg,
                fg=UI_COLORS["text_muted"],
                font=(UI_FONT, 8),
            ).pack(side=tk.LEFT, padx=(4, 0))

            status_label = tk.Label(
                row,
                textvariable=status_var,
                bg="#F1F5F9",
                fg=UI_COLORS["text_secondary"],
                font=(UI_FONT, 8, "bold"),
                padx=9,
                pady=4,
                width=7,
            )
            status_label.grid(row=0, column=4, padx=(0, 8))

            separator = tk.Frame(self.list_inner, bg="#EEF2F7", height=1)
            separator.grid(row=index, column=0, columnspan=5, sticky="sew")

            self.file_rows.append(
                {
                    "path": file_path,
                    "keep_var": keep_var,
                    "keep_entry": keep_entry,
                    "skip_var": skip_var,
                    "skip_entry": skip_entry,
                    "status_var": status_var,
                    "status_label": status_label,
                }
            )

        self.file_count_var.set(f"{len(xlsx_files)} 个文件")
        self.update_status(0, 100, f"已加载 {len(xlsx_files)} 个文件，等待开始")

    def _set_file_status(self, index, text, state="waiting"):
        if index is None or not (0 <= index < len(self.file_rows)):
            return
        row = self.file_rows[index]
        palettes = {
            "waiting": ("#F1F5F9", UI_COLORS["text_secondary"]),
            "running": (UI_COLORS["primary_light"], UI_COLORS["primary"]),
            "success": (UI_COLORS["success_light"], UI_COLORS["success"]),
            "skipped": (UI_COLORS["warning_light"], UI_COLORS["warning"]),
            "failed": (UI_COLORS["danger_light"], UI_COLORS["danger"]),
            "stopped": (UI_COLORS["warning_light"], UI_COLORS["warning"]),
        }
        bg, fg = palettes.get(state, palettes["waiting"])
        row["status_var"].set(text)
        row["status_label"].configure(bg=bg, fg=fg)

    # ────────────────────────────────────────────────────────
    # 任务控制
    # ────────────────────────────────────────────────────────
    def _start(self):
        if not self.file_rows:
            messagebox.showerror("无法开始", "请先选择并读取分类文件夹。")
            return

        self.task_folder = self.folder_var.get().strip()
        self.task_files = [
            {
                "path": row["path"],
                "keep": row["keep_var"].get().strip(),
                "skip": row["skip_var"].get().strip(),
            }
            for row in self.file_rows
        ]

        self.is_running = True
        self.stop_event.clear()
        self.pending_merge = None
        self._set_button_enabled(self.merge_btn, False)
        self.active_file_index = None
        self._set_running_controls(True)

        for index in range(len(self.file_rows)):
            self._set_file_status(index, "等待中", "waiting")

        self._clear_logs()
        self.update_status(0, 100, "正在初始化任务...")
        self.status_dot.configure(fg=UI_COLORS["primary"])
        thread = threading.Thread(target=self._run_task, daemon=True)
        thread.start()

    def _stop(self):
        if self.is_running:
            self.stop_event.set()
            self._set_button_enabled(self.stop_btn, False)
            self.status_dot.configure(fg=UI_COLORS["warning"])
            self.update_status(
                self.progress_var.get(),
                self.progress_bar["maximum"],
                "正在停止并保存当前进度...",
            )
            print("\n[系统] 用户已手动触发停止，等待当前项完成...")

    def _set_running_controls(self, running):
        self._set_button_enabled(self.start_btn, not running)
        self._set_button_enabled(self.stop_btn, running)
        self._set_button_enabled(self.browse_btn, not running)
        self._set_button_enabled(self.scan_btn, not running)
        self.folder_entry.configure(state=tk.DISABLED if running else tk.NORMAL)
        for row in self.file_rows:
            row["keep_entry"].configure(state=tk.DISABLED if running else tk.NORMAL)
            row["skip_entry"].configure(state=tk.DISABLED if running else tk.NORMAL)

    def update_status(self, current, total, text=""):
        safe_total = max(float(total or 0), 0)
        safe_current = max(float(current or 0), 0)

        if safe_total > 0:
            safe_current = min(safe_current, safe_total)
            percent = int(safe_current / safe_total * 100)
            self.progress_bar["maximum"] = safe_total
            self.progress_var.set(safe_current)
            self.progress_percent_var.set(f"{percent}%")
            self.status_var.set(f"{text}  ·  {int(safe_current)}/{int(safe_total)}")
        else:
            self.progress_bar["maximum"] = 100
            self.progress_var.set(0)
            self.progress_percent_var.set("0%")
            self.status_var.set(text)

    def _reset_buttons(self):
        self._set_running_controls(False)
        self.status_dot.configure(
            fg=UI_COLORS["warning"] if self.stop_event.is_set() else UI_COLORS["success"]
        )

    def _do_merge(self):
        if self.is_running:
            messagebox.showwarning("提示", "采集任务进行中，无法合并。")
            return
        if not self.pending_merge:
            messagebox.showwarning("提示", "暂无待合并的采集数据。")
            return
        self._set_button_enabled(self.merge_btn, False)
        threading.Thread(target=self._merge_worker, daemon=True).start()

    def _merge_worker(self):
        pending = self.pending_merge
        try:
            folder = pending["folder"]
            all_results = pending["results"]

            merge_path = os.path.join(
                folder.replace("01INPUT_XLSX", "02OUTPUT_XLSX"),
                f"{os.path.basename(folder)}_合并.xlsx",
            )
            os.makedirs(os.path.dirname(merge_path), exist_ok=True)
            save_xlsx(all_results, merge_path)
            print(f"\n{'=' * 50}")
            print(f"合并完成: {merge_path} ({len(all_results)} 条)")

            # Shopify 转换 + 价格匹配
            print(f"\n{'=' * 50}")
            print("开始 Shopify 转换 + 价格匹配...")
            _styles_to_shopify(merge_path, print)
            print("Shopify 转换 + 价格匹配完成")
            self.pending_merge = None
            self.root.after(0, self.update_status, 0, 0, "合并并转换完成")
        except Exception as exc:
            print(f"合并/转换出错: {exc}")
            import traceback
            traceback.print_exc()
            # 失败允许重试
            self.root.after(0, self._set_button_enabled, self.merge_btn, True)

    # ────────────────────────────────────────────────────────
    # 采集主任务（业务逻辑保持不变，仅补充界面状态）
    # ────────────────────────────────────────────────────────
    def _run_task(self):
        try:
            folder = self.task_folder
            fetch_exchange_rates()

            all_results = []
            grand_success = 0
            grand_fail = 0
            file_count = len(self.task_files)

            for fi, file_info in enumerate(self.task_files):
                if self.stop_event.is_set():
                    break

                self.active_file_index = fi
                self.root.after(0, self._set_file_status, fi, "处理中", "running")

                fpath = file_info["path"]
                fname = os.path.basename(fpath)
                skip_str = file_info["skip"]
                keep_str = file_info.get("keep", "")
                file_skip = None
                file_keep = None
                if keep_str:
                    try:
                        file_keep = []
                        for value in keep_str.replace("，", ",").split(","):
                            value = value.strip().replace("－", "-")
                            if value.lstrip("-").isdigit():
                                file_keep.append(int(value))
                    except Exception:
                        file_keep = None
                if skip_str:
                    try:
                        file_skip = []
                        for value in skip_str.replace("，", ",").split(","):
                            value = value.strip().replace("－", "-")
                            if value.lstrip("-").isdigit():
                                file_skip.append(int(value))
                    except Exception:
                        file_skip = None

                # 输出路径：01INPUT_XLSX → 02OUTPUT_XLSX
                output_file = fpath.replace("01INPUT_XLSX", "02OUTPUT_XLSX")
                os.makedirs(os.path.dirname(output_file), exist_ok=True)

                print(f"\n{'=' * 50}")
                print(f"[{fi + 1}/{file_count}] {fname}")
                print(f"  输出: {output_file}")
                if file_keep:
                    print(f"  保留图片: {file_keep}")
                if file_skip:
                    print(f"  跳过图片: {file_skip}")

                # 读取 + 去重
                targets = read_targets_from_excel(fpath)
                json_url_groups, dup_count = group_targets_by_json_url(targets)
                total = len(json_url_groups)
                print(f"  共 {total} 个唯一商品（跳过 {dup_count} 个重复）")

                if total == 0:
                    print("  无商品，跳过")
                    self.root.after(0, self._set_file_status, fi, "已跳过", "skipped")
                    continue

                # 断点续传
                done_urls = set()
                if os.path.exists(output_file):
                    try:
                        _, existing = load_xlsx(output_file)
                        for row in existing:
                            href = row.get("link-href", "")
                            if href:
                                done_urls.add(str(href).strip())
                        print(f"  断点续传: 已完成 {len(done_urls)} 个")
                    except Exception as exc:
                        print(f"  读取旧文件失败: {exc}")

                todo = {
                    url: group
                    for url, group in json_url_groups.items()
                    if group["original_url"] not in done_urls
                }
                if not todo:
                    print("  全部已完成")
                    if os.path.exists(output_file):
                        try:
                            _, old = load_xlsx(output_file)
                            all_results.extend(old)
                        except Exception:
                            pass
                    self.root.after(0, self._set_file_status, fi, "已完成", "success")
                    continue

                # 加载已有数据
                results = []
                if os.path.exists(output_file):
                    try:
                        _, results = load_xlsx(output_file)
                    except Exception:
                        results = []

                # 启动浏览器
                self.root.after(
                    0,
                    self.update_status,
                    0,
                    len(todo),
                    f"[{fi + 1}/{file_count}] 正在启动浏览器 · {fname}",
                )
                chrome = create_browser()
                tab = chrome.latest_tab
                success_count = 0
                fail_count = 0

                try:
                    for index, (json_url, group) in enumerate(todo.items(), 1):
                        if self.stop_event.is_set():
                            break

                        table_title = group["title"]
                        original_url = group["original_url"]

                        self.root.after(
                            0,
                            self.update_status,
                            index,
                            len(todo),
                            f"[{fi + 1}/{file_count}] {fname} · {table_title[:24]}",
                        )
                        print(f"  [{index}/{len(todo)}] {table_title}")

                        retry_count = 0
                        item = None

                        while retry_count <= MAX_RETRIES:
                            if self.stop_event.is_set():
                                break
                            status, data = fetch_json(tab, json_url)
                            if status == 200 and data:
                                old_skip = globals()["SKIP_POSITIONS"]
                                old_keep = globals()["KEEP_POSITIONS"]
                                globals()["KEEP_POSITIONS"] = file_keep
                                globals()["SKIP_POSITIONS"] = file_skip
                                item = parse_product(data, table_title, original_url)
                                globals()["KEEP_POSITIONS"] = old_keep
                                globals()["SKIP_POSITIONS"] = old_skip
                                break
                            if status == 429:
                                retry_count += 1
                                if retry_count > MAX_RETRIES:
                                    print(f"    429 超过 {MAX_RETRIES} 次，放弃")
                                    break
                                delay = RETRY_BASE_DELAY * retry_count
                                print(f"    429 第{retry_count}次，{delay}秒后重试...")
                                time.sleep(delay)
                            else:
                                retry_count += 1
                                if retry_count > MAX_RETRIES:
                                    print(
                                        f"    失败 {MAX_RETRIES} 次，放弃 "
                                        f"(status={status})"
                                    )
                                    break
                                delay = 5 * retry_count
                                print(f"    失败 (status={status})，{delay}秒后重试...")
                                time.sleep(delay)
                            time.sleep(DELAY)

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

                        time.sleep(DELAY)

                finally:
                    print("  正在关闭浏览器...")
                    chrome.quit()

                if results:
                    save_xlsx(results, output_file)
                    print(f"  文件保存: {output_file} ({len(results)} 条)")

                all_results.extend(results)
                grand_success += success_count
                grand_fail += fail_count

                if self.stop_event.is_set():
                    self.root.after(0, self._set_file_status, fi, "已停止", "stopped")
                    break

                if fail_count > 0 and success_count == 0:
                    self.root.after(0, self._set_file_status, fi, "采集失败", "failed")
                else:
                    self.root.after(0, self._set_file_status, fi, "已完成", "success")

            # 将中断后尚未处理的文件保持为等待状态
            if self.stop_event.is_set() and self.active_file_index is not None:
                for pending_index in range(self.active_file_index + 1, len(self.file_rows)):
                    self.root.after(
                        0,
                        self._set_file_status,
                        pending_index,
                        "未处理",
                        "waiting",
                    )

            # 采集完成，断在这里，等待用户点击「合并并转换」
            if all_results:
                self.pending_merge = {"folder": folder, "results": all_results}
                self.root.after(0, self._set_button_enabled, self.merge_btn, True)
                print(f"\n{'=' * 50}")
                print(
                    f"[等待合并] 采集完成，共 {len(all_results)} 条。\n"
                    "请检查各文件输出后，点击「合并并转换」开始合并 + Shopify转换 + 价格匹配。"
                )
            else:
                print("\n无采集数据，跳过合并")

            final_text = (
                f"总计：{grand_success} 成功，{grand_fail} 失败，"
                f"{len(all_results)} 条数据"
            )
            if self.stop_event.is_set():
                final_text = "任务已中断 · " + final_text
            self.root.after(
                0,
                self.update_status,
                grand_success + grand_fail,
                grand_success + grand_fail,
                final_text,
            )
            print(f"\n{final_text}")

        except Exception as exc:
            print(f"发生异常: {exc}")
            import traceback
            traceback.print_exc()
            if self.active_file_index is not None:
                self.root.after(
                    0,
                    self._set_file_status,
                    self.active_file_index,
                    "发生异常",
                    "failed",
                )
            self.root.after(0, self.update_status, 0, 0, f"任务异常：{exc}")
            self.root.after(0, self.status_dot.configure, {"fg": UI_COLORS["danger"]})

        finally:
            self.is_running = False
            self.root.after(0, self._reset_buttons)


if __name__ == "__main__":
    root = tk.Tk()
    app = ScraperApp(root)
    root.mainloop()
