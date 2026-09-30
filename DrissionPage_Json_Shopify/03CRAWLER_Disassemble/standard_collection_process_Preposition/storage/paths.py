from __future__ import annotations

from pathlib import PurePath
from urllib.parse import urlparse, urlunparse
import re
from ..config import INPUT_DIR_NAME, OUTPUT_DIR_NAME


def to_shopify_json_url(url):
    parsed = urlparse(url)
    path = parsed.path.rstrip("/")
    if not path.endswith(".json"):
        path = f"{path}.json"
    return urlunparse((parsed.scheme, parsed.netloc, path, "", "", ""))


def group_targets_by_json_url(targets):
    json_url_groups = {}
    duplicate_product_count = 0
    for title, url in targets:
        original_url = str(url).strip()
        if not original_url:
            continue
        json_url = to_shopify_json_url(original_url)
        if json_url in json_url_groups:
            duplicate_product_count += 1
            continue
        title_text = str(title).strip() if title is not None else ""
        json_url_groups[json_url] = {
            "title": title_text,
            "original_url": urlunparse(urlparse(original_url)._replace(query="", fragment="")),
        }
    return json_url_groups, duplicate_product_count


def map_to_output_path(path: PurePath) -> PurePath:
    """按基准脚本规则直接替换路径字符串。"""
    return type(path)(str(path).replace(INPUT_DIR_NAME, OUTPUT_DIR_NAME))


def parse_category_pages(raw: str) -> list[tuple[str, str]]:
    """解析每行“标题, URL”的分类列表。"""
    pages: list[tuple[str, str]] = []
    for source_line in raw.splitlines():
        line = source_line.strip()
        if not line or line.startswith("#"):
            continue
        match = re.search(r'(https?://[^\s"\')\]]+)', line)
        if not match:
            continue
        url = match.group(1)
        title = ""
        url_start = line.find(url)
        if url_start > 0:
            title_part = line[:url_start].strip(" \t,(\"'")
            title = re.sub(r'[",)\]]+$', "", title_part).strip()
        if not title:
            title = urlparse(url).path.rstrip("/").split("/")[-1]
        pages.append((title, url))
    return pages
