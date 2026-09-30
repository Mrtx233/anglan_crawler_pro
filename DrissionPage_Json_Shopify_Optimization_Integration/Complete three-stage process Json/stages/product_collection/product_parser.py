# ======================== 产品解析层 ========================
#
# 职责: 把 Shopify 商品 JSON 解析为输出行。含选项排序、变体组合、
#       图片筛选、价格格式化、styles 文本转义、body_html 清洗。
# 依赖: config（SKIP_OPTIONS / 图片过滤全局值 / 汇率缓存）、
#       network_utils（convert_price_to_usd）。
# 禁止: 导入 tkinter、file_utils。
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

from stages.product_collection import config
from stages.product_collection import network_utils

# 便捷别名：与原实现中直接调用 convert_price_to_usd(...) 保持一致
convert_price_to_usd = network_utils.convert_price_to_usd

def find_fallback_variant(product):
    """取商品的主变体：优先 position==1，否则退回第一个；无变体返回空字典。

    当某个选项组合在 variants 里找不到对应变体时，价格等信息用它兜底。
    """
    variants = product.get("variants", [])
    for variant in variants:
        if variant.get("position") == 1:
            return variant
    return variants[0] if variants else {}

def format_combo_value(value):
    """把选项值统一成字符串；None 视为空串，避免拼进 styles 时出现 "None"。"""
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
    """返回规范化并排好序的选项列表，供变体组合与 styles 生成使用。

    处理四件事：
      1. 按 position 排序（原 JSON 顺序不可靠）；
      2. 剔除 skip_options 命中的选项（如 "ships from" 发货地）；
      3. 丢弃无 values 的空选项（单值 option 仍保留，因为它构成变体维度）；
      4. 统一命名——尺码选项一律叫 "Size"，"Type" 改名 "Style"
         （Shopify 保留属性名，直接导入会报错），并把 color 提到首位，
         以便 styles 里第一个选项值就是颜色、能对上按颜色配的图。
    注意：会就地修改传入的 option 字典（改写 name）。
    """
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
    """对各选项的取值做笛卡尔积，返回全部组合（元素是值元组）。

    没有任何选项时返回 [()]（一个空组合），保证调用方至少产出一条变体。
    """
    value_groups = [
        [format_combo_value(value) for value in option.get("values", [])]
        for option in options
    ]
    if not value_groups:
        return [()]
    return list(itertools_product(*value_groups))

def _option_field(options, combo_index):
    """取第 combo_index 个选项对应的 variant 字段名（按真实 position 编号）。"""
    if combo_index < len(options):
        return f"option{options[combo_index].get('position', combo_index + 1)}"
    return f"option{combo_index + 1}"

def _option_field_by_position(options, position):
    """根据 option 的实际 position 返回 variant 字段名。"""
    return f"option{position}"

def variant_matches_combo(variant, combo, options):
    """判断某变体是否与给定选项组合完全一致（逐位比较格式化后的值）。"""
    for index, value in enumerate(combo):
        field = _option_field(options, index)
        if format_combo_value(variant.get(field)) != format_combo_value(value):
            return False
    return True

def find_variant_by_combo(product, combo, options):
    """按完整选项组合查找变体，找不到返回空字典。"""
    for variant in product.get("variants", []):
        if variant_matches_combo(variant, combo, options):
            return variant
    return {}

def find_variant_by_first_option(product, first_option_value, options):
    """只按第一个选项值查找变体（组合对不上时的降级匹配，如颜色相同尺码缺失）。"""
    target = format_combo_value(first_option_value)
    field = _option_field(options, 0)
    for variant in product.get("variants", []):
        if format_combo_value(variant.get(field)) == target:
            return variant
    return {}

def format_price(value, rates=None, currency=None):
    """把价格格式化成紧凑字符串：整数不带小数点，小数最多两位且去尾零。

    一律**按元直接解析**：无论是否含小数点都用 float() 读，"40" 就是 40。
    （旧版曾把不含小数点的值按“分”处理，即 "40" → 0.4；该启发式已移除，
    因为 Shopify 的 price 字段本就是 "40.00" 这类十进制字符串，按分解读
    只会在整数报价的站点上把价格缩小 100 倍。）
    给了 rates 与 currency 时先折算成 USD。无法解析时原样返回。
    """
    if value in (None, ""):
        return ""
    raw = str(value).strip()
    if not raw:
        return ""
    if rates and currency:
        raw = convert_price_to_usd(raw, currency, rates)
    try:
        amount = float(raw)
    except ValueError:
        return raw
    return str(int(amount)) if amount.is_integer() else f"{amount:.2f}".rstrip("0").rstrip(".")

def _is_zero_price(value):
    """判断价格是否等于 0（无法转数字时视为“不是 0”）。"""
    try:
        return float(str(value).strip()) == 0
    except (ValueError, TypeError):
        return False

def format_price2(variant, rates=None):
    """取变体的划线价：compare_at_price 有效时用它，否则退回售价 price。

    compare_at_price 为空串或 0 都视为“没有划线价”，用售价兜底，
    避免出现划线价 0 这种无意义数据。
    """
    currency = variant.get("price_currency", "")
    compare_at_price = variant.get("compare_at_price")
    if compare_at_price and str(compare_at_price).strip() and not _is_zero_price(compare_at_price):
        return format_price(compare_at_price, rates=rates, currency=currency)
    return format_price(variant.get("price", ""), rates=rates, currency=currency)

def format_image_url(url, size="600x600"):
    """在文件名扩展名前插入尺寸标记，得到 CDN 缩略图地址（如 xxx_600x600.jpg）。

    路径中不含扩展名时原样返回；只改 path，query 与 fragment 保留。
    """
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
    """建立 {image_id: 缩略图URL} 映射，供按变体的 image_id 反查图片。

    先按 keep/skip 规则过滤，被过滤掉的图片不会出现在映射里，
    于是引用它的变体也拿不到图片。
    """
    images = _filter_images(product.get("images", []), keep_positions=keep_positions, skip_positions=skip_positions)
    return {
        image.get("id"): format_image_url(image.get("src", ""))
        for image in images
    }

def build_first_image_srcs(product, variant_id=None, limit=None, keep_positions=None, skip_positions=None):
    """按位置顺序取商品图片，拼成 # 分隔的缩略图 URL 串（用于 src_links）。

    指定 variant_id 时优先只取与该变体关联的图；关联图为空则回退全部图片。
    limit 截断数量；keep/skip 规则先于排序结果生效。
    """
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

def build_segment(combo, options, variant, images_by_color_option, rates=None):
    """构造单个变体段：``值1&名称2&值2$售价$划线价@图片``。

    只有第一个选项值不带名称（它的名称在组头里）；选项名与值都做转义，
    避免其中的 & # @ 破坏分段。图片按第一个选项值（通常是颜色）查找，
    查不到就不追加 @ 段。
    """
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
    segment = f"{'&'.join(segment_parts)}${price1}${price2}"
    image_src = images_by_color_option.get(
        format_combo_value(combo[0]) if combo else "", ""
    )
    if image_src:
        segment = f"{segment}@{image_src}"
    return segment

def build_variant_combo(product, keep_positions=None, skip_positions=None, skip_options=None, rates=None):
    """生成完整的 styles1 文本：``组头#变体段#变体段...``。

    对每个选项组合按三级降级找变体：完整组合 → 仅首个选项值 →
    商品主变体，保证任何组合都能拿到一组价格，不会出现空段。
    """
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
    """清洗描述 HTML：丢弃全部属性与链接、剔除 img 与空标签。

    最终只保留纯结构标签（<p>、<ul>、<strong> 等）与转义后的文本，
    以及最多 4 字节的字符（排除 emoji 等补充平面字符，兼容 Excel）。
    """

    void_tags = {
        "area", "base", "br", "col", "embed", "hr", "img", "input",
        "link", "meta", "param", "source", "track", "wbr",
    }

    def __init__(self):
        """初始化输出片段列表与 <a> 嵌套深度计数器。"""
        super().__init__(convert_charrefs=True)
        self.parts = []
        self.anchor_depth = 0

    def handle_starttag(self, tag, attrs):
        """开始标签只写标签名（丢弃属性）；<a> 计数，其内部内容整体忽略。"""
        if tag == "a":
            self.anchor_depth += 1
            return
        if self.anchor_depth or tag == "img":
            return
        self.parts.append(f"<{tag}>")

    def handle_endtag(self, tag):
        """结束标签对称处理；<a> 退出计数，void 标签不补闭合。"""
        if tag == "a":
            self.anchor_depth = max(0, self.anchor_depth - 1)
            return
        if self.anchor_depth or tag in self.void_tags:
            return
        self.parts.append(f"</{tag}>")

    def handle_startendtag(self, tag, attrs):
        """自闭合标签（<br/> 等）同样只写标签名，<a>/<img> 忽略。"""
        if self.anchor_depth or tag in ("a", "img"):
            return
        self.parts.append(f"<{tag}>")

    def handle_data(self, data):
        """文本节点做实体转义，并剔除非 BMP 字符（emoji 等）。"""
        if self.anchor_depth:
            return
        data = "".join(ch for ch in data if ord(ch) <= 0xFFFF)
        self.parts.append(escape(data, quote=False))

    def get_html(self):
        """拼出结果并反复删除空标签（如 <p></p>），直到不再有可删项。"""
        html = "".join(self.parts).strip()
        empty_tag_re = re.compile(r"<([a-zA-Z][a-zA-Z0-9]*)\b[^>]*>\s*</\1>")
        while True:
            cleaned, n = empty_tag_re.subn("", html)
            if n == 0:
                break
            html = cleaned
        return html

def clean_body_html(body_html):
    """清洗商品描述 HTML；空输入返回空串。"""
    if not body_html:
        return ""
    parser = CleanBodyHtmlParser()
    parser.feed(str(body_html))
    parser.close()
    return parser.get_html()

class InvalidProductError(ValueError):
    """响应不是可解析的 Shopify 商品，或商品数据超出可写入范围。"""


def parse_product(data, table_title, original_url):
    """把 Shopify 商品 JSON 解析成一行中间表记录（10 字段）。

    **第二阶段已改用 product_record.parse_product_record 输出 JSON**，
    本函数仅供回退，不再由 scraper_runner 调用。第三、四阶段仍能读取
    它产出的 styles1 中间表。

    图片过滤读 config.KEEP_POSITIONS / SKIP_POSITIONS，调用方需先用
    with_image_filter 按文件/分类规则临时覆盖。src_links 由「按位置取的
    商品图」加「变体专属图（去重后）」组成。
    styles2/styles3 恒为空——所有变体信息都编码在 styles1 里。
    生成结果超出 Excel 单元格上限时抛 InvalidProductError，由调用方跳过该商品。
    """
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
