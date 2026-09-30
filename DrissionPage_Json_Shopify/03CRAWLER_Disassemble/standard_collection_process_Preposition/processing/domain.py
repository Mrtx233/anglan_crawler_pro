from __future__ import annotations

import re

DOMAIN_RE = re.compile(r"^([a-z0-9-]+)\.[a-z]{2,}", re.IGNORECASE)
REPLACE_FIELDS = ("name", "details")
WORD_CHARS = "A-Za-z0-9_-"


def _domain_label(value):
    """取域名主标签：www.thefinespun.com → thefinespun，dailysocks.de → dailysocks。

    不是以域名开头（如纯中文文件夹名）时返回空串，调用方据此跳过替换。
    """
    text = str(value or "").strip()
    if text.lower().startswith("www."):
        text = text[4:]
    match = DOMAIN_RE.match(text)
    return match.group(1).lower() if match else ""


def replace_source_domain(rows, source_label, target_label):
    """把 name / details 中的来源站主标签整体替换为目标站主标签。

    匹配以整个「词」为单位：标签若嵌在更长的词里（Bitterdressmezee、dressmezeeest），
    整个词都替换为目标标签，不做半截替换。大小写不敏感，返回 (rows, 命中处数)，
    命中处数按标签出现次数计。逐行原地修改，不做跨字段联动。
    """
    if not source_label or not target_label or source_label == target_label:
        return rows, 0
    token = re.compile(
        rf"[{WORD_CHARS}]*{re.escape(source_label)}[{WORD_CHARS}]*", re.IGNORECASE
    )
    seen = re.compile(re.escape(source_label), re.IGNORECASE)
    replaced = 0
    for row in rows:
        for field in REPLACE_FIELDS:
            value = row.get(field)
            if not value:
                continue
            text = str(value)
            hits = len(seen.findall(text))
            if not hits:
                continue
            row[field] = token.sub(target_label, text)
            replaced += hits
    return rows, replaced
