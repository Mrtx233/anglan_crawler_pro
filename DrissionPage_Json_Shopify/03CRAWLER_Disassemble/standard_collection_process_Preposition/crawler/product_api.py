from __future__ import annotations

import json


def fetch_json(tab, json_url, log=print):
    tab.listen.start(json_url)
    tab.get(json_url)
    packet = tab.listen.wait(timeout=30)
    status = packet.response.status if packet else 0

    # Shopify 商品不存在时，404 不进行重试
    if status == 404:
        log("  商品不存在 (status=404)，跳过")
        return status, None

    if status == 200:
        text = tab.run_js("return document.body.innerText")
        if not text:
            text = tab.run_js("return document.body.textContent")
        try:
            return status, json.loads(text)
        except json.JSONDecodeError:
            pre_text = tab.run_js("return document.querySelector('pre')?.innerText || ''")
            if pre_text:
                try:
                    return status, json.loads(pre_text)
                except json.JSONDecodeError:
                    pass
            return status, None
    else:
        return status, None
