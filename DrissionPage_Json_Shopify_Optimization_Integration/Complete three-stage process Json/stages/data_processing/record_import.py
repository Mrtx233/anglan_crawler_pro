"""第二阶段的商品记录 JSON → 第三阶段十列中间表。

第二阶段的产物是结构化的 ``{域名}商品.json``（options 数组、variant_images
映射、variant_overrides 覆盖项），而第三阶段的规范与合并只认 ``*商品.xlsx``
中间表。本模块补上这中间的一步：把每条商品记录编码成一行十列数据，
styles 采用「全部维度内嵌进 styles1」的形式，styles2/styles3 留空。

分层约束：本模块属于第三阶段，**不得导入 stages.product_collection**。
转义集来自本阶段的 config._STYLE_RESERVED（与第二阶段 config 中的同名
常量逐字相同），分隔与反转义原语复用 style_utils，因此编码器与解码器
共用同一份「什么算转义」的定义。
"""

import itertools
import math
from pathlib import Path

from stages.data_processing import config
from stages.data_processing.file_utils import read_record_products, save_xlsx

# 第四阶段的 _parse_variant_segment 只读 option1/2/3，第 4 维及以后
# 内嵌进 styles1 会被静默丢弃，因此编码时截断并记录。
MAX_STYLE_DIMENSIONS = 3


class RecordImportError(ValueError):
    """单条商品记录无法编码成中间表行。

    携带商品名与出错字段，便于在整批中止时定位到具体商品。
    """

    def __init__(self, product_name, field, detail):
        self.product_name = product_name
        self.field = field
        super().__init__(f"{product_name} 的 {field} {detail}")


def _style_escape(value):
    """转义保留分隔符（\\ & # @）；先转义反斜杠，避免二次转义。

    与第二阶段 product_parser._style_escape 行为一致，且是
    style_utils._style_unescape 的逆运算。转义集取自本阶段的 config，
    不硬编码字符。
    """
    if value is None:
        return ""
    text = str(value)
    if not any(ch in text for ch in config._STYLE_RESERVED):
        return text
    return (text.replace("\\", "\\\\")
                .replace("&", "\\&")
                .replace("#", "\\#")
                .replace("@", "\\@"))


def _text(value):
    """转成去空白字符串；None 与 float('nan') 都视作空串。"""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return str(value).strip()


def _check_no_dollar(product_name, field, value):
    """拒绝含 ``$`` 的选项名/值。

    ``$`` 既不在转义集内，也不被 style_utils._style_split 当作可转义字符
    （它只把「后跟保留字符的反斜杠」当转义），而解码端用的是裸
    ``str.rsplit("$", 2)``。因此 ``$`` 在原理上无法编码——真值里出现
    ``$`` 会让价格错位。这里选择报错而非静默改写用户数据。
    """
    if "$" in value:
        raise RecordImportError(
            product_name, field,
            f"含无法编码的 $ 符号: {value!r}。styles 格式不支持该字符，"
            f"请先在采集端处理后再转换。",
        )


def _option_dimensions(record, log, keep_options=None):
    """取出参与编码的选项维度，形如 [(名称, [值...]), ...]。

    keep_options 是界面上勾选参与转换的选项名集合；None 表示不限制
    （全部保留）。**未勾选的维度不会被编码进 styles1**——名字、取值、
    变体段都不写，因此第四阶段不会展开它，行数也随之真正减少。

    返回的维度已按 MAX_STYLE_DIMENSIONS 截断，并已校验选项名/值不含 $。
    任一维值为空时整体退回空列表——itertools.product 遇到空序列会产出
    **零个**组合，会让该商品整条消失。
    """
    name = _text(record.get("name")) or "(未命名商品)"
    options = record.get("options")
    if not isinstance(options, list):
        return []
    dimensions = []
    for option in options:
        if not isinstance(option, dict):
            continue
        option_name = _text(option.get("name"))
        if keep_options is not None and option_name not in keep_options:
            continue
        values = option.get("values")
        if not isinstance(values, list):
            values = []
        cleaned = [_text(value) for value in values]
        if not option_name or not cleaned:
            log(f"{name}: 选项 {option_name or '(无名称)'} 缺少可用取值，不参与编码")
            return []
        _check_no_dollar(name, f"选项名 {option_name}", option_name)
        for value in cleaned:
            _check_no_dollar(name, f"选项 {option_name} 的值", value)
        dimensions.append((option_name, cleaned))
    if len(dimensions) > MAX_STYLE_DIMENSIONS:
        dropped = [dimension[0] for dimension in dimensions[MAX_STYLE_DIMENSIONS:]]
        log(f"{name}: 选项 {len(dimensions)} 维，仅编码前 {MAX_STYLE_DIMENSIONS} 维，"
            f"已丢弃 {', '.join(dropped)}")
        dimensions = dimensions[:MAX_STYLE_DIMENSIONS]
    return dimensions


def _override_prices(record, dimension_names):
    """把 variant_overrides 整理成 {(值1, 值2, ...): [售价, 划线价]}。

    键只取**参与编码的那些维度**的值，顺序与 dimension_names 一致，
    因此与 itertools.product 产出的组合元组可比。未勾选的维度不进入键——
    同一维度被剔除后，原本价格不同的两个组合可能落到同一个键上，
    此时保留先出现的那条（JSON 中 overrides 的顺序即第二阶段变体顺序）。
    """
    mapping = {}
    overrides = record.get("variant_overrides")
    if not isinstance(overrides, list):
        return mapping
    for entry in overrides:
        if not isinstance(entry, dict):
            continue
        combo = entry.get("options")
        if not isinstance(combo, dict):
            continue
        key = tuple(_text(combo.get(name)) for name in dimension_names)
        price = entry.get("price")
        if isinstance(price, (list, tuple)) and price and key not in mapping:
            pair = [_text(price[0])]
            pair.append(_text(price[1]) if len(price) > 1 else "")
            mapping[key] = pair
    return mapping


def _image_for(images, first_dimension_name, first_value):
    """按变体段的第 1 维取值查图；**只有颜色维度参与时才给图**。

    variant_images 的键是 ``options[0]`` 的取值，而采集端把 color 排到了
    options 首位（``get_sorted_options``），所以这张映射实质是「颜色值 →
    该颜色的图」。当 Color 被取消勾选、第 1 维换成 Size 之类的维度时，
    仍去查这张映射就是把「尺寸值」当「颜色值」碰运气：查不到倒还好，
    一旦某个尺寸值恰好等于某个颜色值，就会给该变体挂上一张张冠李戴的图。

    因此这里用**维度名**而不是位置来判定：第 1 维叫 color 才追加 ``@``。
    """
    if first_dimension_name.strip().lower() != "color":
        return ""
    return _text(images.get(first_value))


def _build_segment(dimension_names, combo, overrides, image, default_pair):
    """拼一个变体段：``值1&名称2&值2&名称3&值3$售价$划线价[@图片]``。"""
    parts = [_style_escape(combo[0])]
    for index, value in enumerate(combo[1:], start=1):
        parts.append(_style_escape(dimension_names[index]))
        parts.append(_style_escape(value))
    prices = overrides.get(tuple(combo)) or default_pair
    price1 = _text(prices[0]) if prices else ""
    price2 = _text(prices[1]) if len(prices) > 1 else ""
    segment = f"{'&'.join(parts)}${price1}${price2}"
    if image:
        segment = f"{segment}@{image}"
    return segment


def build_styles1(record, log=print, keep_options=None):
    """把一条商品记录编码成内嵌维度的 styles1 文本。

    形如 ``组头#段#段...``；无 options（或全部维度都未勾选）时退化为
    单个无组头的空段，保证每个商品至少产出一行变体。

    keep_options 是界面上勾选参与转换的选项名集合；None 表示不限制。
    **未勾选的维度完全不写进 styles1**，因此第四阶段不会展开它，
    产出行数也随之真正减少。

    价格查找的**已知信息损失**：第二阶段的 variant_overrides 只收录价格
    异于 default_price 的组合，「不在 overrides 里」既可能是「等于默认价」
    也可能是「该组合根本没有变体」。JSON 未保留区分二者所需的信息，
    因此这里统一取 default_price，与「缺失即取默认价」的契约一致。
    """
    dimensions = _option_dimensions(record, log, keep_options=keep_options)

    default_price = record.get("default_price")
    if isinstance(default_price, (list, tuple)) and default_price:
        default_pair = [_text(default_price[0])]
        default_pair.append(_text(default_price[1]) if len(default_price) > 1 else "")
    else:
        default_pair = ["", ""]

    images = record.get("variant_images")
    if not isinstance(images, dict):
        images = {}
    dimension_names = [dimension[0] for dimension in dimensions]
    overrides = _override_prices(record, dimension_names)

    if not dimensions:
        return _build_segment([], ("",), {}, "", default_pair)

    dimension_names = [dimension[0] for dimension in dimensions]
    value_groups = [dimension[1] for dimension in dimensions]

    segments = []
    for combo in itertools.product(*value_groups):
        # 只有第 1 维是颜色时才带图；查不到图不追加 @ 段（而非写空 @），
        # 与第二阶段的处理一致。理由见 _image_for。
        image = _image_for(images, dimension_names[0], combo[0])
        segments.append(_build_segment(dimension_names, combo, overrides, image, default_pair))

    header = _style_escape(dimension_names[0])
    return f"{header}#{'#'.join(segments)}"


def record_to_row(record, log=print, keep_options=None):
    """把一条商品记录转成中间表的十列行。"""
    src_links = record.get("src_links")
    if isinstance(src_links, list):
        images = [_text(url) for url in src_links if _text(url)]
    else:
        images = [_text(src_links)] if _text(src_links) else []
    return {
        "title": _text(record.get("category")),
        "name": _text(record.get("name")),
        "price1": _text((record.get("default_price") or [""])[0]),
        "price2": _text((record.get("default_price") or ["", ""])[1]),
        "styles1": build_styles1(record, log=log, keep_options=keep_options),
        "styles2": "",
        "styles3": "",
        "src_links": "#".join(images),
        "link-href": _text(record.get("link_href")),
        "details": _text(record.get("details")),
    }


def default_output_path(filepath):
    """商品记录 JSON 同目录下的 ``{域名}商品.xlsx``。"""
    source = Path(filepath)
    return source.with_suffix(".xlsx")


def convert_record_file(filepath, output_path=None, log=print, keep_options=None):
    """把单个商品记录 JSON 转成十列中间表，返回输出路径。

    keep_options 是界面上勾选参与转换的选项名；None 表示不限制。
    未勾选的维度不会被写进 styles1。

    整条记录集先全部编码成功、再一次性原子写入，因此任何一条失败都不会
    留下半成品 xlsx。已存在的同名文件会被覆盖（调用方每次运行都重转，
    以免陈旧的中间表被第三阶段照常消费）。
    """
    path = Path(filepath)
    records = read_record_products(path)
    rows = [record_to_row(record, log=log, keep_options=keep_options) for record in records]
    target = Path(output_path) if output_path else default_output_path(path)
    save_xlsx(rows, target)
    log(f"{path.name}: {len(rows)} 条 → {target.name}")
    return target
