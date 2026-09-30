from __future__ import annotations

import hashlib
from ..processing.cells import _clean_cell
from ..resources.columns import WP_HEADERS


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
    values = [
        row.get("Option1 Value", ""),
        row.get("Option2 Value", ""),
        row.get("Option3 Value", ""),
    ]
    values = [value for value in values if value]
    return "-" + ",".join(values) if values else ""


def _csv_join_unique(existing, value):
    value = _clean_cell(value)
    items = [item.strip() for item in _clean_cell(existing).split(",") if item.strip()]
    if value and value not in items:
        items.append(value)
    return ",".join(items)


def build_wp_rows(rows, log=print):
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
                else (
                    row.get("Variant SKU", "")
                    or _generate_parent_sku(handle, used_skus=used_parent_skus)
                )
            )
            parent["Name"] = row.get("Title", "")
            parent["Published"] = "1"
            parent["Visibility in catalog"] = "visible"
            parent["Description"] = row.get("Body (HTML)", "")
            parent["In stock?"] = "1"
            parent["Stock"] = "99999"
            parent["Sale price"] = "" if is_variable else row.get("Variant Price", "")
            parent["Regular price"] = (
                ""
                if is_variable
                else (row.get("Variant Compare At Price", "") or row.get("Variant Price", ""))
            )
            parent["Categories"] = row.get("Collection", "") or row.get("Type", "")
            parent["Images"] = row.get("Image Src", "")
            parent["Parent"] = ""
            parent["Position"] = "0"
            raw_opt1_name = row.get("Option1 Name", "")
            opt1_name = (
                "" if (raw_opt1_name == "Title" or raw_opt1_name.startswith("$")) else raw_opt1_name
            )
            parent["Attribute 1 name"] = opt1_name
            parent["Attribute 1 value(s)"] = (
                ""
                if row.get("Option1 Value", "") == "Default Title"
                else row.get("Option1 Value", "")
            )
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
                "Regular price": row.get("Variant Compare At Price", "")
                or row.get("Variant Price", ""),
                "Categories": parent.get("Categories", ""),
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

    return [
        {header: item.get(header, "") for header in WP_HEADERS}
        for items in products.values()
        for item in items
    ]
