from __future__ import annotations

from ..config import MAX_STYLES_LENGTH
from ..processing.currency import format_price, format_price2
from ..processing.html_cleaner import clean_body_html
from ..processing.images import (
    build_first_image_srcs,
    prepare_data_for_image_skips,
    build_images_by_first_option,
)
from ..processing.options import find_fallback_variant, get_sorted_options
from ..processing.variants import build_variant_combo


class InvalidProductError(ValueError):
    """响应不是可解析的 Shopify 商品，或商品数据超出可写入范围。"""


def validate_product_json(data):
    """验证 Shopify 商品 JSON 的基本结构；空响应和错误页不能标为成功。"""
    if not isinstance(data, dict) or not isinstance(data.get("product"), dict):
        raise InvalidProductError("JSON 缺少 product 对象")
    product = data["product"]
    if (
        not isinstance(product.get("id"), int)
        or isinstance(product["id"], bool)
        or product["id"] <= 0
    ):
        raise InvalidProductError("商品缺少有效 id")
    for field in ("title", "handle"):
        if not isinstance(product.get(field), str) or not product[field].strip():
            raise InvalidProductError(f"商品缺少有效 {field}")
    variants = product.get("variants")
    if not isinstance(variants, list) or not variants:
        raise InvalidProductError("商品缺少非空 variants 列表")
    for variant in variants:
        if (
            not isinstance(variant, dict)
            or not isinstance(variant.get("id"), int)
            or isinstance(variant["id"], bool)
            or variant["id"] <= 0
        ):
            raise InvalidProductError("商品变体缺少有效 id")
        try:
            price = float(variant.get("price"))
        except (ValueError, TypeError):
            raise InvalidProductError("商品变体缺少有效 price") from None
        if not 0 <= price < float("inf") or isinstance(variant.get("price"), bool):
            raise InvalidProductError("商品变体 price 无效")
    for field in ("options", "images"):
        entries = product.get(field, [])
        if not isinstance(entries, list) or any(not isinstance(entry, dict) for entry in entries):
            raise InvalidProductError(f"商品 {field} 必须是对象列表")
    return product


def parse_product(
    data,
    table_title,
    original_url,
    *,
    skip_positions=None,
    keep_positions=None,
    skip_options=None,
    rates=None,
    raw_skip_spec="",
):
    validate_product_json(data)
    prepared_data, skip_positions = prepare_data_for_image_skips(
        data,
        skip_positions,
        raw_skip_spec,
    )
    product = prepared_data["product"]
    variant = find_fallback_variant(product)
    options = get_sorted_options(product, skip_options=skip_options)
    src_links = [
        url
        for url in build_first_image_srcs(
            product, keep_positions=keep_positions, skip_positions=skip_positions
        ).split("#")
        if url
    ]
    for url in build_images_by_first_option(
        product, options, keep_positions=keep_positions, skip_positions=skip_positions
    ).values():
        if url and url not in src_links:
            src_links.append(url)
    styles1 = build_variant_combo(
        product,
        skip_positions=skip_positions,
        keep_positions=keep_positions,
        skip_options=skip_options,
        rates=rates,
    )
    if len(styles1) > MAX_STYLES_LENGTH:
        raise InvalidProductError(
            f"styles1 长度 {len(styles1)} 超过上限 {MAX_STYLES_LENGTH}，已跳过"
        )
    return {
        "title": table_title,
        "name": product.get("title", ""),
        "price1": format_price(
            variant.get("price", ""),
            rates=rates,
            currency=variant.get("price_currency", ""),
        ),
        "price2": format_price2(variant, rates=rates),
        "styles1": styles1,
        "styles2": "",
        "styles3": "",
        "src_links": "#".join(src_links),
        "link-href": original_url,
        "details": clean_body_html(product.get("body_html", "")),
    }
