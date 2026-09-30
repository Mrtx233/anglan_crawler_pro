from __future__ import annotations

from itertools import product as itertools_product


def find_fallback_variant(product):
    variants = product.get("variants", [])
    for variant in variants:
        if variant.get("position") == 1:
            return variant
    return variants[0] if variants else {}


def format_combo_value(value):
    return "" if value is None else str(value)


def get_sorted_options(product, skip_options=None):
    options = sorted(
        product.get("options", []),
        key=lambda option: option.get("position") or 0,
    )
    if skip_options:
        skip_lower = [s.lower() for s in skip_options]
        options = [
            opt
            for opt in options
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
        (i for i, opt in enumerate(multi_value) if (opt.get("name") or "").lower() == "color"),
        None,
    )
    if color_idx is not None and color_idx > 0:
        color_opt = multi_value.pop(color_idx)
        multi_value.insert(0, color_opt)
    return multi_value


def build_option_combos(options):
    value_groups = [
        [format_combo_value(value) for value in option.get("values", [])] for option in options
    ]
    if not value_groups:
        return [()]
    return list(itertools_product(*value_groups))


def _option_field(options, combo_index):
    if combo_index < len(options):
        return f"option{options[combo_index].get('position', combo_index + 1)}"
    return f"option{combo_index + 1}"


SIZE_VALUES = {
    "xs",
    "s",
    "m",
    "l",
    "xl",
    "xxl",
    "xxxl",
    "xxxxl",
    "2xl",
    "3xl",
    "4xl",
    "5xl",
    "xs/s",
    "s/m",
    "m/l",
    "l/xl",
    "one size",
    "os",
    "free size",
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
