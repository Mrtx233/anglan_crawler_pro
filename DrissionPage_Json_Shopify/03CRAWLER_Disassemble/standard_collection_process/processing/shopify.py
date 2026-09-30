from __future__ import annotations

from itertools import product as itertools_product
import pandas as pd
import random
import re
import string
from ..processing.cells import _clean_cell
from .style_codec import _style_split, _style_split_first, _style_unescape
from ..resources.columns import SHOPIFY_COLUMNS


def _clean_handle_name(name):
    name = _clean_cell(name)
    cleaned = re.sub(r"[^a-zA-Z\s]", "", name).lower()
    return cleaned.replace(" ", "-").strip("-")


def _generate_unique_handle(base, used_handles):
    while True:
        digits = "".join(random.choices(string.digits, k=3))
        letters1 = "".join(random.choices(string.ascii_lowercase, k=2))
        letters2 = "".join(random.choices(string.ascii_lowercase, k=2))
        handle = (
            f"{base}-{letters1}{digits}{letters2}" if base else f"-{letters1}{digits}{letters2}"
        )
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


def build_shopify_frame(df):
    required = [
        "title",
        "name",
        "price1",
        "price2",
        "details",
        "src_links",
        "styles1",
        "styles2",
        "styles3",
    ]
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
        images = [
            img.strip() for img in _clean_cell(item.get("src_links", "")).split("#") if img.strip()
        ]
        combinations = list(
            itertools_product(
                option1_values or [""], option2_values or [""], option3_values or [""]
            )
        )
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
                row["Variant Compare At Price"] = segment_info["price2"] or _clean_cell(
                    item.get("price2", "")
                )
                row["Variant Requires Shipping"] = "TRUE"
                row["Variant Taxable"] = "FALSE"
                row["Variant Image"] = segment_info["image"]

            if index < len(images):
                row["Image Src"] = images[index]
                row["Image Position"] = index + 1

            rows.append(row)

    return pd.DataFrame(rows, columns=SHOPIFY_COLUMNS)
