"""Shopify JSON 商品变体解析工具 — 从 product JSON 构建 styles 编码字符串。"""

import json
import urllib.request
from html import escape as _html_escape
from html.parser import HTMLParser
from itertools import product as _cartesian
from urllib.parse import parse_qs, urlparse, urlunparse


# ── URL 转换 ───────────────────────────────────────────────

def to_shopify_json_url(url):
    """将任意 Shopify 商品 URL 转换为 /products/xxx.json 接口地址。"""
    parts = urlparse(url)
    path = parts.path.rstrip("/")
    if not path.endswith(".json"):
        path += ".json"
    return urlunparse((parts.scheme, parts.netloc, path, "", "", ""))


def _extract_variant_id(url):
    """从 ?variant=123 参数中取出数字 ID。"""
    vals = parse_qs(urlparse(url).query).get("variant", [])
    if vals:
        try:
            return int(vals[0])
        except ValueError:
            pass
    return None


def group_targets_by_json_url(targets):
    """
    去重: 多条记录指向同一商品 JSON 时只保留首次出现的标题。
    返回 (json_url_groups dict, duplicate_count int)。
    """
    groups = {}
    dupes = 0
    for title, raw_url in targets:
        url = str(raw_url).strip()
        if not url:
            continue
        key = to_shopify_json_url(url)
        if key in groups:
            dupes += 1
            continue
        groups[key] = {
            "title": str(title).strip() if title is not None else "",
            "original_url": urlunparse(urlparse(url)._replace(query="", fragment="")),
        }
    return groups, dupes


# ── 汇率 ───────────────────────────────────────────────────

_rates_cache: dict = {}


def fetch_exchange_rates():
    """从 open.er-api.com 拉取以 USD 为基准的汇率表（带内存缓存）。"""
    global _rates_cache
    if _rates_cache:
        return _rates_cache
    try:
        req = urllib.request.Request(
            "https://open.er-api.com/v6/latest/USD",
            headers={"User-Agent": "Mozilla/5.0"},
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            _rates_cache = json.loads(resp.read().decode()).get("rates", {})
        print(f"[rates] 已获取 {len(_rates_cache)} 条汇率")
    except Exception as exc:
        print(f"[rates] 获取失败: {exc}")
    return _rates_cache


def _to_usd(amount_str, currency, rates):
    """把 amount_str 按 currency→USD 换算，失败则原样返回。"""
    if not amount_str or not currency or not rates:
        return amount_str
    cur = str(currency).upper().strip()
    if cur == "USD":
        return amount_str
    rate = rates.get(cur)
    if not rate:
        return amount_str
    try:
        return f"{float(str(amount_str).strip()) / rate:.2f}"
    except (ValueError, TypeError):
        return amount_str


# ── 变体查找 ────────────────────────────────────────────────

def _by_id(product, vid):
    for v in product.get("variants", []):
        if v.get("id") == vid:
            return v
    return {}


def find_fallback_variant(product):
    """position=1 的 variant 优先，否则取第一个。"""
    variants = product.get("variants", [])
    for v in variants:
        if v.get("position") == 1:
            return v
    return variants[0] if variants else {}


# ── Options 处理 ────────────────────────────────────────────

def _sorted_options(product, skip_options=None):
    """按 position 排序，过滤掉单值 option 和 skip_options 中的选项；Color 优先提到首位。"""
    opts = sorted(product.get("options", []), key=lambda o: o.get("position") or 0)
    if skip_options:
        skip_low = [s.lower() for s in skip_options]
        opts = [o for o in opts if not any(s in (o.get("name") or "").lower() for s in skip_low)]
    multi = [o for o in opts if len(o.get("values") or []) > 1]
    ci = next((i for i, o in enumerate(multi) if (o.get("name") or "").lower() == "color"), None)
    if ci and ci > 0:
        multi.insert(0, multi.pop(ci))
    return multi


# ── 价格格式化 ──────────────────────────────────────────────

def _fmt_price(raw, rates=None, currency=None):
    """把原始价格字符串格式化为简洁数字，支持汇率转换和分→元。"""
    if raw in (None, ""):
        return ""
    s = str(raw).strip()
    if not s:
        return ""
    if rates and currency:
        s = _to_usd(s, currency, rates)
    if "." in s:
        try:
            val = float(s)
        except ValueError:
            return s
    else:
        try:
            val = int(s) / 100
        except ValueError:
            return s
    if val.is_integer():
        return str(int(val))
    return f"{val:.2f}".rstrip("0").rstrip(".")


def _fmt_price2(variant, rates=None):
    """compare_at_price 优先，否则用 price。"""
    cur = variant.get("price_currency", "")
    cap = variant.get("compare_at_price")
    if cap and str(cap).strip():
        return _fmt_price(cap, rates=rates, currency=cur)
    return _fmt_price(variant.get("price", ""), rates=rates, currency=cur)


# ── 图片处理 ────────────────────────────────────────────────

def _resize_url(url, size="600x600"):
    """在 Shopify 图片 URL 扩展名前插入尺寸后缀。"""
    if not url:
        return ""
    p = urlparse(url)
    path, dot, slash = p.path, p.path.rfind("."), p.path.rfind("/")
    if dot > slash:
        path = f"{path[:dot]}_{size}{path[dot:]}"
    return urlunparse((p.scheme, p.netloc, path, p.params, p.query, p.fragment))


def _images_by_id(product, skip_positions=None):
    imgs = product.get("images", [])
    if skip_positions:
        imgs = [i for i in imgs if (i.get("position") or 0) not in skip_positions]
    return {i.get("id"): _resize_url(i.get("src", "")) for i in imgs}


def build_first_image_srcs(product, variant_id=None, limit=None, skip_positions=None):
    """取商品前 N 张图片（或 variant 关联图片），返回 # 拼接的 URL 串。"""
    imgs = sorted(product.get("images", []), key=lambda i: i.get("position") or 0)
    if skip_positions:
        imgs = [i for i in imgs if (i.get("position") or 0) not in skip_positions]
    if variant_id is not None:
        matched = [i for i in imgs if variant_id in (i.get("variant_ids") or [])]
        if matched:
            imgs = matched
    chosen = imgs if limit is None else imgs[:limit]
    return "#".join(_resize_url(i.get("src", "")) for i in chosen if i.get("src"))


# ── 笛卡尔积 + variant 匹配 ─────────────────────────────────

def _opt_field(options, idx):
    """过滤后 option 列表的第 idx 项对应 variant 的 optionN 字段名。"""
    if idx < len(options):
        return f"option{options[idx].get('position', idx + 1)}"
    return f"option{idx + 1}"


def _combo_values(options):
    groups = [[str(v) if v is not None else "" for v in o.get("values", [])] for o in options]
    return list(_cartesian(*groups)) if groups else [()]


def _match_variant(product, combo, options):
    for v in product.get("variants", []):
        if all(
            str(v.get(_opt_field(options, i)) or "") == str(val)
            for i, val in enumerate(combo)
        ):
            return v
    return {}


def _match_by_first(product, first_val, options):
    target = str(first_val)
    field = _opt_field(options, 0)
    for v in product.get("variants", []):
        if str(v.get(field) or "") == target:
            return v
    return {}


# ── 图片按首 option 分组 ─────────────────────────────────────

def _images_by_first_opt(product, options, skip_positions=None):
    img_map = _images_by_id(product, skip_positions=skip_positions)
    field = _opt_field(options, 0)
    result = {}
    for v in product.get("variants", []):
        val = str(v.get(field) or "")
        if val and val not in result:
            iid = v.get("image_id")
            if iid and iid in img_map:
                result[val] = img_map[iid]
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


# ── 构建 styles 编码 ────────────────────────────────────────

def _build_segment(combo, options, variant, img_by_first, rates=None):
    parts = []
    if combo:
        parts.append(_style_escape(str(combo[0])))
        for i, val in enumerate(combo[1:], 1):
            parts.extend([_style_escape(str(options[i].get("name", ""))),
                          _style_escape(str(val))])

    cur = variant.get("price_currency", "")
    p1 = _fmt_price(variant.get("price"), rates=rates, currency=cur)
    p2 = _fmt_price(variant.get("compare_at_price") or variant.get("price"), rates=rates, currency=cur)
    first_val = str(combo[0]) if combo else ""
    seg = f"{'&'.join(parts)}${p1}${p2}"
    img = img_by_first.get(first_val, "")
    if img:
        seg += f"@{img}"
    return seg


def build_variant_combo(product, skip_positions=None, skip_options=None, rates=None):
    """
    核心函数：从 product JSON 构建 styles1 编码字符串。
    格式: OptionName#val1&SubName&SubVal$p1$p2@img#val2&...
    """
    options = _sorted_options(product, skip_options=skip_options)
    combos = _combo_values(options)
    fallback = find_fallback_variant(product)
    img_by_first = _images_by_first_opt(product, options, skip_positions=skip_positions)

    segments = []
    for combo in combos:
        v = _match_variant(product, combo, options)
        if not v:
            v = _match_by_first(product, combo[0], options) if combo else {}
        if not v:
            v = fallback
        segments.append(_build_segment(combo, options, v, img_by_first, rates=rates))

    header = str(options[0].get("name", "")) if options else ""
    return f"{_style_escape(header)}#{'#'.join(segments)}" if header else "#".join(segments)


# ── HTML 清洗 ───────────────────────────────────────────────

_VOID_TAGS = frozenset({
    "area", "base", "br", "col", "embed", "hr", "img",
    "input", "link", "meta", "param", "source", "track", "wbr",
})


class _HtmlCleaner(HTMLParser):
    """去除 <a> 标签，保留 <img src>，过滤高位 Unicode。"""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self._buf: list[str] = []

    def handle_starttag(self, tag, attrs):
        if tag == "a":
            return
        if tag == "img":
            src = dict(attrs).get("src", "")
            self._buf.append(f'<img src="{_html_escape(src, quote=True)}">' if src else "<img>")
        else:
            self._buf.append(f"<{tag}>")

    def handle_endtag(self, tag):
        if tag == "a" or tag in _VOID_TAGS:
            return
        self._buf.append(f"</{tag}>")

    def handle_startendtag(self, tag, attrs):
        if tag == "a":
            return
        if tag == "img":
            src = dict(attrs).get("src", "")
            self._buf.append(f'<img src="{_html_escape(src, quote=True)}">' if src else "<img>")
        else:
            self._buf.append(f"<{tag}>")

    def handle_data(self, data):
        cleaned = "".join(ch for ch in data if ord(ch) <= 0xFFFF)
        self._buf.append(_html_escape(cleaned, quote=False))

    def result(self):
        return "".join(self._buf).strip()


def clean_body_html(html_text):
    if not html_text:
        return ""
    parser = _HtmlCleaner()
    parser.feed(str(html_text))
    parser.close()
    return parser.result()
