"""将 pipeline 输出的 styles XLSX 转换为 Shopify 商品导入格式（XLSX + CSV），并执行价格匹配。"""

import random
import re
import string
from itertools import product as _cartesian
from pathlib import Path

import pandas as pd


# ── Shopify 列定义 ──────────────────────────────────────────

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

_REQUIRED_INPUT_COLS = [
    "title", "name", "price1", "price2",
    "details", "src_links", "styles1", "styles2", "styles3",
]

# ── 预设价格库（用于匹配） ────────────────────────────────────

PRICE_LIBRARY = sorted([
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
])


# ── Handle / SKU 生成 ───────────────────────────────────────

def _rand_handle(base, used):
    """生成随机 handle: base-XnnnnYY 格式，确保唯一。"""
    while True:
        d = "".join(random.choices(string.digits, k=3))
        l1 = "".join(random.choices(string.ascii_lowercase, k=2))
        l2 = "".join(random.choices(string.ascii_lowercase, k=2))
        h = f"{base}-{l1}{d}{l2}" if base else f"-{l1}{d}{l2}"
        if h not in used:
            used.add(h)
            return h


def _rand_sku(used):
    """生成随机 SKU: 123AB-C45DE-678AB 格式，确保唯一。"""
    while True:
        d1 = random.choice("123456789")
        d2 = "".join(random.choices(string.digits, k=2))
        l1 = "".join(random.choices(string.ascii_uppercase, k=2))
        l2 = random.choice(string.ascii_uppercase)
        d3 = "".join(random.choices(string.digits, k=2))
        l3 = "".join(random.choices(string.ascii_uppercase, k=2))
        d4 = "".join(random.choices(string.digits, k=3))
        sku = f"{d1}{d2}{l1}-{l2}{d3}{l3}-{d4}{l1}"
        if sku not in used:
            used.add(sku)
            return sku


# ── 单元格清洗 ──────────────────────────────────────────────

def _cell(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return str(val).strip()


def _clean_handle(name):
    cleaned = re.sub(r"[^a-zA-Z\s]", "", _cell(name)).lower()
    return cleaned.replace(" ", "-").strip("-")


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


# ── styles 编码解析 ──────────────────────────────────────────

def _parse_group(value):
    """把 styles 字段拆成 (option_name, [segments])。"""
    parts = [p.strip() for p in _style_split(_cell(value), "#") if p.strip()]
    if not parts:
        return "", [""]
    if len(parts) == 1:
        return _style_unescape(parts[0]), [""]
    return _style_unescape(parts[0]), parts[1:]


def _parse_segment(seg):
    """解析单个 variant segment: option值 + 价格 + 图片。"""
    text = _cell(seg)
    image = ""
    text, image, has_image = _style_split_first(text, "@")
    if has_image:
        image = image.strip()
    p1, p2 = "", ""
    price_parts = text.rsplit("$", 2)
    if len(price_parts) == 3:
        text, p1, p2 = price_parts
    tokens = _style_split(text, "&") if text else []

    def _field(index):
        return _style_unescape(tokens[index]) if len(tokens) > index else ""

    return {
        "opt1_val": _field(0),
        "opt2_name": _field(1),
        "opt2_val": _field(2),
        "opt3_name": _field(3),
        "opt3_val": _field(4),
        "price1": p1,
        "price2": p2,
        "image": image,
    }


# ── 价格匹配 ────────────────────────────────────────────────

def _match_prices(shopify_path):
    """
    对 Shopify XLSX 执行价格匹配：
    - 在 75%~125% 区间内从 PRICE_LIBRARY 选取最近价格
    - 同一价格不重复分配
    - 输出 _Shopify_价格匹配 + _Shopify_未匹配到的价格 文件
    """
    path = Path(shopify_path)
    df = pd.read_excel(path).reset_index(drop=True)
    df["匹配价格"] = None

    pool = list(PRICE_LIBRARY)
    used = set()

    for handle, group in df.groupby("Handle"):
        if pd.isna(handle):
            continue
        valid = []
        for idx, row in group.iterrows():
            p = row.get("Variant Price")
            if pd.isna(p):
                continue
            try:
                valid.append((idx, float(p)))
            except (ValueError, TypeError):
                continue
        if not valid:
            continue

        available = [p for p in pool if p not in used]
        for idx, orig in sorted(valid, key=lambda x: x[1]):
            lo, hi = orig * 0.75, orig * 1.25
            candidates = [p for p in available if lo <= p <= hi]
            if candidates:
                picked = candidates[0]
                used.add(picked)
                available.remove(picked)
                df.at[idx, "匹配价格"] = picked

    # 清理文件名中的 _Shopify_原价 后缀
    stem = path.stem
    for suffix in ("_Shopify_原价",):
        if stem.endswith(suffix):
            stem = stem[: -len(suffix)]
            break

    xlsx_dir = path.parent
    csv_dir = xlsx_dir.parent / "csv"
    csv_dir.mkdir(exist_ok=True)

    unused = [p for p in pool if p not in used]
    unused_df = pd.DataFrame({"未匹配价格": unused})

    matched_xlsx = xlsx_dir / f"{stem}_Shopify_价格匹配.xlsx"
    matched_csv = csv_dir / f"{stem}_Shopify_价格匹配.csv"
    unused_xlsx = xlsx_dir / f"{stem}_Shopify_未匹配到的价格.xlsx"
    unused_csv = csv_dir / f"{stem}_Shopify_未匹配到的价格.csv"

    df.to_excel(matched_xlsx, index=False)
    df.to_csv(matched_csv, index=False, encoding="utf-8-sig")
    unused_df.to_excel(unused_xlsx, index=False)
    unused_df.to_csv(unused_csv, index=False, encoding="utf-8-sig")

    print(f"价格匹配: {matched_xlsx}")
    print(f"未匹配: {unused_xlsx}")
    print(f"已用价格: {len(used)}, 剩余: {len(unused)}")

    return matched_xlsx


# ── 主转换入口 ───────────────────────────────────────────────

def convert_to_shopify(input_file_name):
    """
    读取 pipeline 输出的 styles XLSX，展开变体笛卡尔积，
    生成 Shopify 格式的 XLSX + CSV，然后执行价格匹配。
    返回 (price_xlsx_path, matched_xlsx_path)。
    """
    inp = Path(input_file_name)
    xlsx_dir = inp.parent / "xlsx"
    csv_dir = inp.parent / "csv"
    xlsx_dir.mkdir(exist_ok=True)
    csv_dir.mkdir(exist_ok=True)

    out_xlsx = xlsx_dir / f"{inp.stem}_Shopify_原价.xlsx"
    out_csv = csv_dir / f"{inp.stem}_Shopify_原价.csv"

    src_df = pd.read_excel(inp)
    missing = [c for c in _REQUIRED_INPUT_COLS if c not in src_df.columns]
    if missing:
        raise ValueError(f"输入文件缺少必要列: {', '.join(missing)}")

    handles_used = set()
    skus_used = set()
    rows = []

    for _, item in src_df.iterrows():
        handle = _rand_handle(_clean_handle(item.get("name", "")), handles_used)

        opt1_name, opt1_segs = _parse_group(item.get("styles1", ""))
        opt2_name, opt2_segs = _parse_group(item.get("styles2", ""))
        opt3_name, opt3_segs = _parse_group(item.get("styles3", ""))
        images = [s.strip() for s in _cell(item.get("src_links", "")).split("#") if s.strip()]
        combos = list(_cartesian(opt1_segs or [""], opt2_segs or [""], opt3_segs or [""]))
        total = max(len(combos), len(images), 1)

        for i in range(total):
            row = {col: "" for col in SHOPIFY_COLUMNS}
            row["Link-Href"] = _cell(item.get("link-href", ""))
            row["Handle"] = handle
            row["Variant Weight Unit"] = "kg"

            # 首行写商品信息
            if i == 0:
                row["Title"] = _cell(item.get("name", ""))
                row["Body (HTML)"] = _cell(item.get("details", ""))
                row["Collection"] = _cell(item.get("title", ""))
                row["Vendor"] = _cell(item.get("title", ""))
                row["Published"] = "TRUE"
                row["Gift Card"] = "FALSE"
                row["SEO Title"] = row["Title"]
                row["SEO Description"] = row["Title"]

            # 变体信息
            if i < len(combos):
                seg = _parse_segment(combos[i][0])
                row["Option1 Name"] = opt1_name
                row["Option1 Value"] = seg["opt1_val"] or _cell(combos[i][0])
                row["Option2 Name"] = seg["opt2_name"] or opt2_name
                row["Option2 Value"] = seg["opt2_val"] or _cell(combos[i][1])
                row["Option3 Name"] = seg["opt3_name"] or opt3_name
                row["Option3 Value"] = seg["opt3_val"] or _cell(combos[i][2])
                row["Variant SKU"] = _rand_sku(skus_used)
                row["Variant Grams"] = 0
                row["Variant Inventory Tracker"] = "shopify"
                row["Variant Inventory Qty"] = 999
                row["Variant Inventory Policy"] = "deny"
                row["Variant Fulfillment Service"] = "manual"
                row["Variant Price"] = seg["price1"] or _cell(item.get("price1", ""))
                row["Variant Compare At Price"] = seg["price2"] or _cell(item.get("price2", ""))
                row["Variant Requires Shipping"] = "TRUE"
                row["Variant Taxable"] = "FALSE"
                row["Variant Image"] = seg["image"]

            # 图片信息
            if i < len(images):
                row["Image Src"] = images[i]
                row["Image Position"] = i + 1

            rows.append(row)

    result = pd.DataFrame(rows, columns=SHOPIFY_COLUMNS)
    result.to_excel(out_xlsx, index=False)
    result.to_csv(out_csv, index=False, encoding="utf-8-sig")

    print(f"Shopify XLSX: {out_xlsx}")
    print(f"Shopify CSV: {out_csv}")
    print(f"商品数: {len(src_df)}，输出行数: {len(rows)}")

    matched = _match_prices(str(out_xlsx))
    return out_xlsx, matched
