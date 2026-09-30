# ======================== 转换流水线层 ========================
#
# 职责: 合并结果 -> Shopify 格式 -> 价格匹配 -> WooCommerce 格式。
# 依赖: config（PRICE_LIBRARY / SHOPIFY_COLUMNS / WP_HEADERS）、
#       style_utils（styles 文本解析函数）。
# 禁止: 导入 tkinter、network_utils、file_utils、main。
#
# 说明: 函数名保留原有下划线前缀，便于与原实现逐一对照；
#       对外入口为 styles_to_shopify（= _styles_to_shopify）。
# ==============================================================

import csv
import hashlib
import random
import re
import string
from itertools import product as itertools_product
from pathlib import Path

import pandas as pd

import config
from style_utils import (
    _style_split,
    _style_split_first,
    _style_unescape,
)

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
    """将 styles xlsx 转换为 Shopify 格式并自动价格匹配。返回原价 Shopify 文件路径。"""
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
            row = {col: "" for col in config.SHOPIFY_COLUMNS}
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
    result_df = pd.DataFrame(rows, columns=config.SHOPIFY_COLUMNS)
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
    _shopify_to_wp(str(csv_path), log)
    if price_match_path:
        # 匹配价 CSV 在 csv/，不能在 xlsx/ 中仅替换扩展名查找。
        price_csv = csv_dir / f"{Path(price_match_path).stem}.csv"
        _shopify_to_wp(str(price_csv), log)
    log("Shopify 转 WP 完成")

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
    price_library_sorted = sorted(config.PRICE_LIBRARY)

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
    wp_headers_no_tags = [h for h in config.WP_HEADERS if h != "Tags"]
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


# ======================== 对外入口 ========================
# 保留原名以兼容既有调用；同时提供无下划线的公开别名。
styles_to_shopify = _styles_to_shopify
price_match = _price_match
shopify_to_wp = _shopify_to_wp
