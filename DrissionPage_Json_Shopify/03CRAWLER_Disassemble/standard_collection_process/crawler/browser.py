from __future__ import annotations

from DrissionPage import Chromium, ChromiumOptions
import random
from ..config import PROXY_URL


def create_browser(log=print):
    co = ChromiumOptions()
    co.incognito(on_off=True)
    co.no_imgs(True)
    co.set_load_mode("normal")
    # 固定使用本地代理 127.0.0.1:7897
    co.set_argument(f"--proxy-server={PROXY_URL}")
    random_port = random.randint(9222, 9322)
    co.set_local_port(random_port)
    for attempt in range(3):
        try:
            chrome = Chromium(co)
            log(f"浏览器启动成功，端口: {random_port}")
            return chrome
        except Exception as e:
            log(f"端口 {random_port} 启动失败 (尝试 {attempt + 1}/3): {e}")
            random_port = random.randint(9222, 9322)
            co.set_local_port(random_port)
    raise RuntimeError("无法启动浏览器，已重试多次")
