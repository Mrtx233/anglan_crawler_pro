from __future__ import annotations

import re
from html import escape
from html.parser import HTMLParser


class CleanBodyHtmlParser(HTMLParser):
    void_tags = {
        "area",
        "base",
        "br",
        "col",
        "embed",
        "hr",
        "img",
        "input",
        "link",
        "meta",
        "param",
        "source",
        "track",
        "wbr",
    }

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []

    def handle_starttag(self, tag, attrs):
        if tag in ("a", "img"):
            return
        self.parts.append(f"<{tag}>")

    def handle_endtag(self, tag):
        if tag == "a" or tag in self.void_tags:
            return
        self.parts.append(f"</{tag}>")

    def handle_startendtag(self, tag, attrs):
        if tag in ("a", "img"):
            return
        self.parts.append(f"<{tag}>")

    def handle_data(self, data):
        data = "".join(ch for ch in data if ord(ch) <= 0xFFFF)
        self.parts.append(escape(data, quote=False))

    def get_html(self):
        html = "".join(self.parts).strip()
        empty_tag_re = re.compile(r"<([a-zA-Z][a-zA-Z0-9]*)\b[^>]*>\s*</\1>")
        while True:
            cleaned, n = empty_tag_re.subn("", html)
            if n == 0:
                break
            html = cleaned
        return html


def clean_body_html(body_html):
    if not body_html:
        return ""
    parser = CleanBodyHtmlParser()
    parser.feed(str(body_html))
    parser.close()
    return parser.get_html()
