"""商品中间表 styles 字段的分隔与反转义。"""

import config

def _style_split(text, sep):
    """按未转义的 sep 拆分，保留转义序列本身（供后续按层还原）。"""
    parts = []
    buf = []
    i, n = 0, len(text)
    while i < n:
        ch = text[i]
        if ch == "\\" and i + 1 < n and text[i + 1] in config._STYLE_RESERVED:
            buf.append(text[i:i + 2])
            i += 2
            continue
        if ch == sep:
            parts.append("".join(buf))
            buf = []
            i += 1
            continue
        buf.append(ch)
        i += 1
    parts.append("".join(buf))
    return parts

def _style_split_first(text, sep):
    """在第一个未转义的 sep 处切一刀，返回 (前段, 后段, 是否找到)。"""
    i, n = 0, len(text)
    while i < n:
        ch = text[i]
        if ch == "\\" and i + 1 < n and text[i + 1] in config._STYLE_RESERVED:
            i += 2
            continue
        if ch == sep:
            return text[:i], text[i + 1:], True
        i += 1
    return text, "", False

def _style_unescape(text):
    """还原 _style_escape 转义过的字符。"""
    if "\\" not in text:
        return text
    out = []
    i, n = 0, len(text)
    while i < n:
        ch = text[i]
        if ch == "\\" and i + 1 < n and text[i + 1] in config._STYLE_RESERVED:
            out.append(text[i + 1])
            i += 2
            continue
        out.append(ch)
        i += 1
    return "".join(out)
