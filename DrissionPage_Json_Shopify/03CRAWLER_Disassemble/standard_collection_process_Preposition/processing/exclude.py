from __future__ import annotations


def split_by_keywords(rows, keywords, field="name"):
    """剔除 field 包含任一关键词的商品行。

    关键词按整体字面量匹配，不拆词、不加词边界：条目里的空格也是要匹配的字符，
    所以 "Gift Card" 命中 "Shop Vista Gift Card"，不命中 "Cardigan"。
    大小写不敏感。每个条目互相独立，增删不影响其他条目。

    返回 (保留的行, [(原值, [命中的关键词, ...]), ...])。
    """
    needles = [(word, str(word).strip()) for word in (keywords or [])]
    needles = [(word, text) for word, text in needles if text]
    if not needles:
        return rows, []
    lowered = [(word, text.lower()) for word, text in needles]

    kept = []
    removed = []
    for row in rows:
        value = row.get(field)
        text = str(value) if value else ""
        haystack = text.lower()
        hits = [word for word, needle in lowered if needle in haystack]
        if hits:
            removed.append((text, hits))
        else:
            kept.append(row)
    return kept, removed
