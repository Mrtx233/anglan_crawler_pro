# ======================== 产品解析层 ========================
#
# 职责: 把 Shopify 商品 JSON 解析为输出行。含选项排序、变体组合、
#       图片筛选、价格格式化、styles 文本转义、body_html 清洗。
# 依赖: config（SKIP_OPTIONS / 图片过滤全局值 / 汇率缓存）、
#       network_utils（convert_price_to_usd）。
# 禁止: 导入 tkinter、converter、file_utils。
#
# 图片过滤全局值:
#   parse_product 读取 config.KEEP_POSITIONS / config.SKIP_POSITIONS，
#   与原实现读取模块级全局完全等价。
#   main.py 通过 with_image_filter(keep, skip) 临时切换，退出时自动还原
#   （原实现用 try/finally + globals()[...] 手工保存还原）。
# ============================================================

import re
from contextlib import contextmanager
from html import escape
from html.parser import HTMLParser
from itertools import product as itertools_product
from urllib.parse import urlparse, urlunparse

import config
import network_utils

# 便捷别名：与原实现中直接调用 convert_price_to_usd(...) 保持一致
convert_price_to_usd = network_utils.convert_price_to_usd

def find_fallback_variant(product):
    variants = product.get("variants", [])
    for variant in variants:
        if variant.get("position") == 1:
            return variant
    return variants[0] if variants else {}

def format_combo_value(value):
    return "" if value is None else str(value)

def _is_size_option(option):
    """判断 option 是否为尺码选项。名称含 'size' 或 values 都是尺码值。"""
    name = (option.get("name") or "").lower().strip()
    if "size" in name:
        return True
    values = [str(v).lower().strip() for v in (option.get("values") or [])]
    if values and all(v in config.SIZE_VALUES for v in values):
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
    multi_value = [opt for opt in options if opt.get("values")]  # 保留单值 option 的变体维度
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

def _style_escape(value):
    """转义选项值/选项名里的保留分隔符（\\ & # @）；不含则原样返回。"""
    if value is None:
        return ""
    text = str(value)
    if not any(ch in text for ch in config._STYLE_RESERVED):
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
        if ch == "\\" and i + 1 < n and text[i + 1] in config._STYLE_RESERVED:
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
        if ch == "\\" and i + 1 < n and text[i + 1] in config._STYLE_RESERVED:
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
        if ch == "\\" and i + 1 < n and text[i + 1] in config._STYLE_RESERVED:
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

class InvalidProductError(ValueError):
    """响应不是可解析的 Shopify 商品，或商品数据超出可写入范围。"""


def parse_product(data, table_title, original_url):
    product = data.get("product", {})
    variant = find_fallback_variant(product)
    options = get_sorted_options(product, skip_options=config.SKIP_OPTIONS)
    variant_image_urls = [
        url for url in build_images_by_color_option(
            product, options, keep_positions=config.KEEP_POSITIONS, skip_positions=config.SKIP_POSITIONS
        ).values() if url
    ]
    src_links = [
        src for src in build_first_image_srcs(
            product, keep_positions=config.KEEP_POSITIONS, skip_positions=config.SKIP_POSITIONS
        ).split("#") if src
    ]
    for url in variant_image_urls:
        if url not in src_links:
            src_links.append(url)
    styles1 = build_variant_combo(
        product,
        keep_positions=config.KEEP_POSITIONS,
        skip_positions=config.SKIP_POSITIONS,
        skip_options=config.SKIP_OPTIONS,
        rates=config._exchange_rates_cache,
    )
    if len(styles1) > config.MAX_STYLES_LENGTH:
        raise InvalidProductError(
            f"styles1 长度 {len(styles1)} 超过上限 {config.MAX_STYLES_LENGTH}，已跳过"
        )
    return {
        "title": table_title,
        "name": product.get("title", ""),
        "price1": format_price(
            variant.get("price", ""),
            rates=config._exchange_rates_cache,
            currency=variant.get("price_currency", ""),
        ),
        "price2": format_price2(variant, rates=config._exchange_rates_cache),
        "styles1": styles1,
        "styles2": "",
        "styles3": "",
        "src_links": "#".join(src_links),
        "link-href": original_url,
        "details": clean_body_html(product.get("body_html", "")),
    }


# ======================== 图片过滤临时切换 ========================
@contextmanager
def with_image_filter(keep_positions, skip_positions):
    """在 with 块内临时设置 config.KEEP_POSITIONS / SKIP_POSITIONS。

    等价于原实现的:
        old_keep = globals()["KEEP_POSITIONS"]
        old_skip = globals()["SKIP_POSITIONS"]
        globals()["KEEP_POSITIONS"] = file_keep
        globals()["SKIP_POSITIONS"] = file_skip
        try:
            item = parse_product(...)
        finally:
            globals()["KEEP_POSITIONS"] = old_keep
            globals()["SKIP_POSITIONS"] = old_skip

    异常时同样保证还原（contextmanager 的 finally 语义）。
    """
    old_keep = config.KEEP_POSITIONS
    old_skip = config.SKIP_POSITIONS
    config.KEEP_POSITIONS = keep_positions
    config.SKIP_POSITIONS = skip_positions
    try:
        yield
    finally:
        config.KEEP_POSITIONS = old_keep
        config.SKIP_POSITIONS = old_skip
