# ======================== 浏览器工具层 ========================
#
# 职责: 封装 DrissionPage 的浏览器生命周期与 JS 执行细节。
# 依赖: config（参数）、logger（日志）、DrissionPage。
# 禁止: 导入 tkinter、collector、file_utils。
#
# 说明: create_browser 的代理/端口/禁图三项在原实现中硬编码，
#       此处改为从 config 读取并支持参数覆盖，启动行为保持一致。
# ========================================================

import json
import random
import time

from stages.link_collection import config
from stages.link_collection import logger

# 未安装 DrissionPage 时仍保证本模块可导入，便于纯函数单测；
# 真正调用 create_browser() 时才抛 ImportError。
try:
    from DrissionPage import Chromium, ChromiumOptions
    from DrissionPage.errors import ContextLostError

    DRISSIONPAGE_AVAILABLE = True
except ImportError:  # pragma: no cover - 取决于运行环境
    Chromium = None
    ChromiumOptions = None
    DRISSIONPAGE_AVAILABLE = False

    class ContextLostError(Exception):
        """占位异常类，保证 except ContextLostError 语法可用。"""


def create_browser(
    proxy=config.PROXY_SERVER,
    port_range=config.PORT_RANGE,
    disable_images=config.DISABLE_IMAGES,
    incognito=config.USE_INCOGNITO,
    retries=config.BROWSER_START_RETRIES,
):
    """启动 Chromium 并返回实例。

    端口在 port_range 内随机取值，启动失败换端口重试 retries 次。
    proxy 为空/None 时不传 --proxy-server（直连）。
    全部重试失败抛 RuntimeError。
    """
    co = ChromiumOptions()
    if incognito:
        co.incognito(on_off=True)

    # 只采集 HTML/JSON/链接，关闭图片加载可明显减少列表页等待时间。
    if disable_images:
        co.set_argument("--blink-settings=imagesEnabled=false")

    if proxy:
        co.set_argument(f"--proxy-server={proxy}")

    low, high = port_range
    random_port = random.randint(low, high)
    co.set_local_port(random_port)

    for attempt in range(retries):
        try:
            chrome = Chromium(co)
            logger.log(f"浏览器启动成功，端口: {random_port}")
            return chrome
        except Exception as e:
            logger.log(
                f"端口 {random_port} 启动失败 (尝试 {attempt + 1}/{retries}): {e}"
            )
            random_port = random.randint(low, high)
            co.set_local_port(random_port)

    raise RuntimeError("无法启动浏览器，已重试多次")


def interrupt_page(tab) -> None:
    """尽力中止正在加载的页面。

    不同 DrissionPage 版本可能没有 stop_loading 接口，缺失时安全忽略。
    供停止按钮在采集线程之外调用，任何异常都不向上抛。
    """
    if tab is None:
        return
    try:
        stop_loading = getattr(tab, "stop_loading", None)
        if callable(stop_loading):
            stop_loading()
    except Exception:
        pass


def run_js_with_retry(
    tab,
    script,
    default=None,
    retries=config.JS_DEFAULT_RETRIES,
    stop_event=None,
):
    """执行 JavaScript；重试等待可被停止事件立即打断。

    - ContextLostError: 间隔 JS_CONTEXT_LOST_DELAY 后重试
    - 其他异常: 间隔 JS_OTHER_ERROR_DELAY 后重试，末次失败打印日志
    - 返回 None 时统一替换为 default
    """
    for attempt in range(retries + 1):
        if stop_event is not None and stop_event.is_set():
            return default
        try:
            result = tab.run_js(script)
            return default if result is None else result
        except ContextLostError:
            if attempt >= retries:
                return default
            delay = config.JS_CONTEXT_LOST_DELAY
        except Exception as e:
            if attempt >= retries:
                logger.log(f"  JavaScript 执行失败: {e}")
                return default
            delay = config.JS_OTHER_ERROR_DELAY

        if stop_event is not None:
            if stop_event.wait(delay):
                return default
        else:
            time.sleep(delay)


def get_outer_html(tab, stop_event=None) -> str:
    """读取整页 HTML 文本；失败返回空字符串。"""
    return run_js_with_retry(
        tab,
        "return document.documentElement.outerHTML;",
        default="",
        stop_event=stop_event,
    )


def evaluate_xpath(tab, xpath, stop_event=None):
    """在页面内用 document.evaluate 求值 XPath，返回原始结果列表。

    - 正常: list[str]
    - XPath 语法/求值错误: dict {"__xpath_error__": "..."}（调用方负责抛错）
    - JS 执行彻底失败: default []
    """
    return run_js_with_retry(
        tab,
        f"""
        return (() => {{
            const xpath = {json.dumps(xpath)};
            const results = [];
            try {{
                const snapshot = document.evaluate(
                    xpath, document, null,
                    XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null
                );
                for (let i = 0; i < snapshot.snapshotLength; i++) {{
                    const node = snapshot.snapshotItem(i);
                    const val = node.nodeType === 2 ? node.value : (node.href || node.textContent || '');
                    if (val) results.push(val.trim());
                }}
            }} catch(e) {{
                return {{__xpath_error__: String(e)}};
            }}
            return results;
        }})();
        """,
        default=[],
        stop_event=stop_event,
    )
