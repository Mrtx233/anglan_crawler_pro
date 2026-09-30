from __future__ import annotations

from urllib.parse import urlparse, urlunparse
from ..processing.options import _option_field, format_combo_value


def format_image_url(url, size="600x600"):
    if not url:
        return ""
    parsed = urlparse(url)
    path = parsed.path
    dot_index = path.rfind(".")
    slash_index = path.rfind("/")
    if dot_index > slash_index:
        path = f"{path[:dot_index]}_{size}{path[dot_index:]}"
    return urlunparse(
        (
            parsed.scheme,
            parsed.netloc,
            path,
            parsed.params,
            parsed.query,
            parsed.fragment,
        )
    )


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
    images = _filter_images(
        product.get("images", []), keep_positions=keep_positions, skip_positions=skip_positions
    )
    return {image.get("id"): format_image_url(image.get("src", "")) for image in images}


def build_first_image_srcs(
    product, variant_id=None, limit=None, keep_positions=None, skip_positions=None
):
    images = sorted(
        product.get("images", []),
        key=lambda image: image.get("position") or 0,
    )
    images = _filter_images(images, keep_positions=keep_positions, skip_positions=skip_positions)
    if variant_id is not None:
        variant_images = [
            image for image in images if variant_id in (image.get("variant_ids") or [])
        ]
        if variant_images:
            images = variant_images
    selected_images = images if limit is None else images[:limit]
    srcs = [format_image_url(image.get("src", "")) for image in selected_images if image.get("src")]
    return "#".join(srcs)


def build_images_by_first_option(product, options, keep_positions=None, skip_positions=None):
    """按主选项建立 {选项值: 图片URL} 映射。

    优先用名为 color 的选项（get_sorted_options 已将其排到最前）；
    没有 color 时退回第一个选项（如 Device/Size），让单选项商品的
    每个 variant 也能带上自己的 image_id 对应图。
    """
    images_by_id = build_images_by_id(
        product, keep_positions=keep_positions, skip_positions=skip_positions
    )
    color_opt = next((opt for opt in options if (opt.get("name") or "").lower() == "color"), None)
    if not color_opt:
        color_opt = options[0] if options else None
    if not color_opt:
        return {}
    field = f"option{color_opt.get('position', 1)}"
    result = {}
    for variant in product.get("variants", []):
        opt_value = format_combo_value(variant.get(field))
        if opt_value and opt_value not in result:
            image_id = variant.get("image_id")
            if image_id and image_id in images_by_id:
                result[opt_value] = images_by_id[image_id]
    return result


def parse_image_positions(spec):
    if not str(spec or "").strip():
        return None
    positions = []
    for token in str(spec).replace("，", ",").replace("－", "-").split(","):
        try:
            positions.append(int(token.strip()))
        except ValueError:
            continue
    return positions


def prepare_data_for_image_skips(data, configured_positions, raw_skip_spec):
    """保留原数据；负数统一由图片过滤器按 position 倒数解析。"""
    positions = (
        parse_image_positions(raw_skip_spec)
        if str(raw_skip_spec or "").strip()
        else configured_positions
    )
    return data, positions
