# ======================== 浏览器工具层 ========================
#
# 职责: 封装 DrissionPage 浏览器创建。
# 依赖: DrissionPage。
# 禁止: 导入 tkinter、GUI。
#
# 说明: 代理地址与端口区间在原实现中硬编码于函数体内，
#       此处提取为模块级常量，行为与默认值保持不变。
# ==============================================================

import random

from DrissionPage import Chromium, ChromiumOptions

# 本地代理地址；设为 None 或空字符串表示直连（不传 --proxy-server）。
# 原实现硬编码为 "http://127.0.0.1:7897"，此处保持同值。
PROXY_SERVER = "http://127.0.0.1:7897"

# 本地调试端口随机区间（原实现: random.randint(9222, 9322)）
PORT_RANGE = (9222, 9322)

# 启动失败后的换端口重试次数
START_RETRIES = 3

def create_browser():
    co = ChromiumOptions()
    co.incognito(on_off=True)
    # 代理设置，如有需要自行取消或修改
    if PROXY_SERVER:
        co.set_argument(f"--proxy-server={PROXY_SERVER}")
    low, high = PORT_RANGE
    random_port = random.randint(low, high)
    co.set_local_port(random_port)
    for attempt in range(START_RETRIES):
        try:
            chrome = Chromium(co)
            print(f"浏览器启动成功，端口: {random_port}")
            return chrome
        except Exception as e:
            print(f"端口 {random_port} 启动失败 (尝试 {attempt + 1}/{START_RETRIES}): {e}")
            random_port = random.randint(low, high)
            co.set_local_port(random_port)
    raise RuntimeError("无法启动浏览器，已重试多次")
