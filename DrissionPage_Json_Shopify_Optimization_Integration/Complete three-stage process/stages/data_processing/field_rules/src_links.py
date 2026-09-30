"""保留 src_links 独有的前四张，再追加与 styles1 共有的图片。"""
import math

from stages.data_processing.style_utils import _style_split, _style_split_first


SRC_ONLY_LIMIT = 4


def _text(value):
    """把单元格值转成去空白字符串；None 与 float('nan') 都视作空串。"""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ''
    return str(value).strip()


def extract_images_from_styles1(value):
    """识别未转义的 #/@ 分隔符，不将选项中的转义字符当成图片。"""
    images = set()
    for segment in _style_split(_text(value), '#'):
        _, url, found = _style_split_first(segment, '@')
        if found and url.strip():
            images.add(url.strip())
    return images


def parse_src_links(value):
    """把 src_links 按 # 拆成图片列表，返回 (有序列表, 去重集合)。

    src_links 由采集端直接以 # 拼接、未做转义，故此处用裸 split 即可；
    调用方主要用列表长度做前后对比统计。
    """
    images = [url.strip() for url in _text(value).split('#') if url.strip()]
    return images, set(images)


def normalize_src_links(value, styles1):
    """沿用原脚本排序和重复项语义；不补入仅在 styles1 出现的图片。"""
    variant_images = extract_images_from_styles1(styles1)
    images, _ = parse_src_links(value)
    only_src = [url for url in images if url not in variant_images]
    shared = [url for url in images if url in variant_images]
    return '#'.join(only_src[:SRC_ONLY_LIMIT] + shared)
