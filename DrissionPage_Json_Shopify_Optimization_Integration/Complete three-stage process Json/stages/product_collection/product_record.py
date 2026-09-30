# ======================== 商品记录层 ========================
#
# 职责: 把 Shopify 商品 JSON 解析成第二阶段的输出记录（结构化字典），
#       供 file_utils 落盘为 {域名}商品.json。
# 依赖: config（全局图片过滤值 / SKIP_OPTIONS / 汇率缓存）、
#       product_parser（选项规范化、变体查找、图片映射、价格格式化）。
# 禁止: 导入 tkinter、file_utils、data_processing 任何模块。
#
# 输出字段:
#   category / name / link_href / details / options / default_price /
#   variant_images / variant_overrides / src_links
#
# 图片过滤全局值:
#   与 product_parser.parse_product 一致，读 config.KEEP_POSITIONS /
#   config.SKIP_POSITIONS，调用方需先用 product_parser.with_image_filter
#   按文件/分类规则临时覆盖。
#
# 语义约定（消费方需知）:
#   1. options 里的选项名是**规范化后**的名字——尺码类统一为 "Size"、
#      原 "Type" 改为 "Style"（Shopify 保留属性名）、color 被提到首位。
#      原始 Shopify 选项名不可从本记录还原。这与第四阶段写进 Shopify 的
#      Option N Name 一致。
#   2. variant_images 的键即 options[0].values 的取值，无论 options[0]
#      叫什么名字（有 Color 时是颜色，否则可能是尺寸等）。这是 styles1
#      时代就存在的隐含约定，此处只是显式化。
#   3. variant_overrides 只收录价格**不同于 default_price** 的选项组合，
#      相同的省略。因此「缺失」必须被理解为「该组合取 default_price」；
#      该字段为空即代表全部变体同价。
#   4. src_links 是数组，元素为缩略图 URL；与 styles1 时代同一批图，
#      只是不再拼成 # 分隔的字符串。
#   5. 含 config.EXCLUDED_VALUE_KEYWORDS 关键字的选项值（如 "Custom Size"）
#      在产出时即被剔除，options / variant_overrides / variant_images 三处
#      保持一致，不会出现某处残留。被剔空的选项整个丢弃。
# ============================================================

from stages.product_collection import config
from stages.product_collection import product_parser


def _price_pair(variant, rates):
    """取某个变体的 (售价, 划线价) 字符串对。

    与 product_parser.parse_product 用同一对函数，保证价格格式与旧输出一致。
    """
    return [
        product_parser.format_price(
            variant.get("price", ""),
            rates=rates,
            currency=variant.get("price_currency", ""),
        ),
        product_parser.format_price2(variant, rates=rates),
    ]


def _resolve_variant(product, combo, options, fallback_variant):
    """按三级降级为某个选项组合找变体：完整组合 → 仅首个选项值 → 主变体。

    与 product_parser.build_variant_combo 的降级顺序完全一致，保证任何
    组合都能拿到一组价格，不会出现空价格。
    """
    variant = product_parser.find_variant_by_combo(product, combo, options)
    if not variant:
        variant = product_parser.find_variant_by_first_option(
            product, combo[0], options
        ) if combo else {}
    if not variant:
        variant = fallback_variant
    return variant


def _is_excluded_value(value):
    """判断选项值是否应被剔除。

    "Custom Size" 这类定制项不是真实的可售尺码，导出到 Shopify 会变成
    一个无法履约的变体，因此在产出记录时就丢掉。按子串匹配（大小写与
    首尾空格不敏感），一并覆盖 "Custom Made"、"Custom Order" 等变体写法。
    """
    return any(word in str(value or "").strip().lower() for word in config.EXCLUDED_VALUE_KEYWORDS)


def _filter_option_values(options):
    """剔除含排除关键字的选项值；被剔空的选项整个丢掉（没有取值就没有变体维度）。

    返回新的选项字典，不改动传入的对象——下游的 options、variant_overrides、
    variant_images 都从这一份派生，过滤放在这里三处自然保持一致。
    """
    filtered = []
    for option in options:
        values = [
            value for value in (option.get("values") or [])
            if not _is_excluded_value(value)
        ]
        if not values:
            continue
        filtered.append({**option, "values": values})
    return filtered


def build_options_payload(options):
    """把规范化后的选项列表转成 [{"name": ..., "values": [...]}, ...]。"""
    return [
        {
            "name": product_parser.format_combo_value(option.get("name", "")),
            "values": [
                product_parser.format_combo_value(value)
                for value in (option.get("values") or [])
            ],
        }
        for option in options
    ]


def build_variant_overrides(product, options, default_price, rates):
    """生成逐变体价格覆盖项：只收录价格异于 default_price 的组合。

    每条形如 {"options": {"Color": "Navy", "Size": "M"}, "price": ["42","48"]}，
    键用 options 里的真实（规范化后）选项名。价格与 default_price 相同的
    组合直接省略，让常见商品（全部同价）产出空数组。
    """
    overrides = []
    fallback_variant = product_parser.find_fallback_variant(product)
    for combo in product_parser.build_option_combos(options):
        variant = _resolve_variant(product, combo, options, fallback_variant)
        price = _price_pair(variant, rates)
        if price == default_price:
            continue
        combo_map = {}
        for index, value in enumerate(combo):
            option_name = product_parser.format_combo_value(
                options[index].get("name", "")
            )
            combo_map[option_name] = product_parser.format_combo_value(value)
        overrides.append({"options": combo_map, "price": price})
    return overrides


def build_src_links(product, variant_image_urls):
    """按位置取商品图，再补上各变体专属图（去重），返回 URL 数组。

    推导逻辑与旧 parse_product 的 src_links 完全一致，只是不再拼成
    # 分隔字符串。注意不 import data_processing.field_rules.src_links——
    那会跨层，且该模块的规则属于第三阶段。
    """
    src_links = [
        src for src in product_parser.build_first_image_srcs(
            product,
            keep_positions=config.KEEP_POSITIONS,
            skip_positions=config.SKIP_POSITIONS,
        ).split("#") if src
    ]
    for url in variant_image_urls:
        if url not in src_links:
            src_links.append(url)
    return src_links


def parse_product_record(data, category, original_url):
    """把一个 Shopify 商品 JSON 解析成第二阶段的输出记录。

    签名与 product_parser.parse_product 对齐，便于 scraper_runner 直接替换。
    图片过滤读 config.KEEP_POSITIONS / SKIP_POSITIONS（由 with_image_filter 覆盖）。

    data 里没有 product 键时抛 InvalidProductError：某些站点在 200 响应里
    返回登录页或错误 JSON，若不拦住会写出一条全空记录。
    """
    product = data.get("product")
    if not isinstance(product, dict) or not product:
        raise product_parser.InvalidProductError("响应中没有 product 数据，已跳过")

    rates = config._exchange_rates_cache
    # get_sorted_options 会就地改写 option 的 name，每个商品只调用一次，
    # 后续步骤共用这一个列表，避免二次改名。剔除定制项后再派生后续字段，
    # 保证 options / variant_overrides / variant_images 三者一致。
    options = _filter_option_values(product_parser.get_sorted_options(
        product, skip_options=config.SKIP_OPTIONS
    ))

    variant_override_map = product_parser.build_images_by_color_option(
        product,
        options,
        keep_positions=config.KEEP_POSITIONS,
        skip_positions=config.SKIP_POSITIONS,
    )
    variant_image_urls = [url for url in variant_override_map.values() if url]

    fallback_variant = product_parser.find_fallback_variant(product)
    default_price = _price_pair(fallback_variant, rates)

    return {
        "category": category,
        "name": product.get("title", ""),
        "link_href": original_url,
        "details": product_parser.clean_body_html(product.get("body_html", "")),
        "options": build_options_payload(options),
        "default_price": default_price,
        "variant_images": dict(variant_override_map),
        "variant_overrides": build_variant_overrides(
            product, options, default_price, rates
        ),
        "src_links": build_src_links(product, variant_image_urls),
    }
