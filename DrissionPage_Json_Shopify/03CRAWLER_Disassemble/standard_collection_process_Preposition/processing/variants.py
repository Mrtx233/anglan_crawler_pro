from __future__ import annotations

from ..processing.currency import format_price, format_price2
from .style_codec import _style_escape
from ..processing.images import build_images_by_first_option
from ..processing.options import (
    _option_field,
    build_option_combos,
    find_fallback_variant,
    format_combo_value,
    get_sorted_options,
)


def build_segment(combo, options, variant, images_by_color_option, rates=None):
    segment_parts = []
    if combo:
        segment_parts.append(_style_escape(format_combo_value(combo[0])))
        for index, value in enumerate(combo[1:], start=1):
            option_name = format_combo_value(options[index].get("name", ""))
            segment_parts.extend(
                [
                    _style_escape(option_name),
                    _style_escape(format_combo_value(value)),
                ]
            )
    currency = variant.get("price_currency", "")
    price1 = format_price(variant.get("price"), rates=rates, currency=currency)
    price2 = format_price2(variant, rates=rates)
    color_value = format_combo_value(combo[0]) if combo else ""
    image_src = images_by_color_option.get(color_value, "")
    segment = f"{'&'.join(segment_parts)}${price1}${price2}"
    if image_src:
        segment = f"{segment}@{image_src}"
    return segment


def build_variant_combo(
    product, skip_positions=None, skip_options=None, rates=None, keep_positions=None
):
    options = get_sorted_options(product, skip_options=skip_options)
    combos = build_option_combos(options)
    fallback_variant = find_fallback_variant(product)
    images_by_first_opt = build_images_by_first_option(
        product, options, skip_positions=skip_positions, keep_positions=keep_positions
    )
    fields = [_option_field(options, i) for i in range(len(options))]
    exact_variants = {}
    first_variants = {}
    for variant in product.get("variants", []):
        key = tuple(format_combo_value(variant.get(field)) for field in fields)
        exact_variants.setdefault(key, variant)
        if key:
            first_variants.setdefault(key[0], variant)
    segments = []
    for combo in combos:
        variant = (
            exact_variants.get(combo)
            or (first_variants.get(combo[0]) if combo else None)
            or fallback_variant
        )
        segments.append(build_segment(combo, options, variant, images_by_first_opt, rates=rates))
    first_option_name = format_combo_value(options[0].get("name", "")) if options else ""
    if first_option_name:
        return f"{_style_escape(first_option_name)}#{'#'.join(segments)}"
    return "#".join(segments)
