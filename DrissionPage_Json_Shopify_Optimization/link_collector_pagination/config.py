# ======================== 配置层 ========================
#
# 职责: 集中管理全部常量与任务输入结构。
# 依赖: 仅标准库（本模块是依赖图的叶子，禁止导入其他业务模块）。
#
# 说明:
#   - 采集参数、浏览器参数、输出文件名均从原 link_collector_pagination.py
#     中硬编码的位置提取而来，行为保持完全一致。
#   - CollectConfig 是 GUI 与采集端之间唯一的契约：GUI 负责组装，
#     collector 负责消费，两边互不感知。
# ========================================================

from dataclasses import dataclass, field
from typing import List, Optional

# ======================== UI 样式 ========================
UI_COLORS = {
    "background": "#F3F5F9",        # 窗口底色
    "card": "#FFFFFF",              # 卡片底色
    "card_alt": "#F8FAFC",          # 输入框 / 文本框底色
    "border": "#DEE4ED",
    "text": "#18263B",
    "muted": "#64748B",
    "primary": "#2563EB",
    "primary_hover": "#1D4ED8",
    "primary_disabled": "#CBD5E1",
    "danger": "#B91C1C",
    "progress_trough": "#E8EEF7",
    "log_background": "#142033",
    "log_text": "#DCE6F3",
}

# 界面字体：按顺序取第一个系统已安装的候选，都没有时用 fallback
UI_FONT_CANDIDATES = ("PingFang SC", "Microsoft YaHei UI", "Noto Sans CJK SC")
UI_FONT_FALLBACK = "TkDefaultFont"
MONO_FONT_CANDIDATES = ("Menlo", "Consolas")
MONO_FONT_FALLBACK = "Consolas"

# 窗口尺寸
WINDOW_GEOMETRY = "1120x820"
WINDOW_MIN_SIZE = (940, 740)
APP_TITLE = "Shopify 商品链接采集"   # 窗口标题栏
APP_HEADING = "商品链接采集"         # 界面内的大标题

# 日志独立窗口
LOG_WINDOW_TITLE = "运行日志 · 商品链接采集"
LOG_WINDOW_GEOMETRY = "900x560"
LOG_WINDOW_MIN_SIZE = (600, 360)

# 界面文案
CATEGORY_PLACEHOLDER = (
    "每行一个分类：礼服, https://example.com/collections/dresses"
)
OPTIONS_HINT = "留空：不限页数、自动识别链接"
FOOTER_HINT = "结果按站点保存为 Excel，自动合并已有链接"
STATUS_READY = "就绪 · 填写分类后开始采集"
STATUS_STARTING = "正在启动浏览器…"
STATUS_STOPPING = "正在停止并保存…"

# ======================== 采集参数 ========================
# 首页无链接时的重试次数与间隔
MAX_RETRIES = 3
RETRY_DELAY = 2.0

# tab.get() 完成后额外等待页面稳定的时间
PAGE_SETTLE_WAIT = 0.35

# 页面上下文丢失（ContextLostError）后快速重试一次的等待时间
CONTEXT_LOST_RETRY_DELAY = 0.8

# JavaScript 执行失败的重试间隔
JS_CONTEXT_LOST_DELAY = 0.6
JS_OTHER_ERROR_DELAY = 0.3
JS_DEFAULT_RETRIES = 2

# ======================== 浏览器参数 ========================
# 本地代理地址。设为 None 或空字符串表示直连（不传 --proxy-server）。
# 原实现硬编码在 create_browser() 内部，此处提取为可配置项。
PROXY_SERVER = "http://127.0.0.1:7897"

# 本地调试端口随机区间（原实现: random.randint(9222, 9322)）
PORT_RANGE = (9222, 9322)

# 浏览器启动失败后的换端口重试次数
BROWSER_START_RETRIES = 3

# 关闭图片加载：只采集 HTML/JSON/链接，可明显减少列表页等待时间
DISABLE_IMAGES = True

# 使用隐身模式
USE_INCOGNITO = True

# ======================== 输出 ========================
# 分类索引文件名（域名 → 分类标题 → 分类URL）
LINK_INDEX_FILENAME = "links_index.json"

# 恢复文件名中的标记（旧 XLSX 读取失败时另存，绝不覆盖原文件）
RECOVERY_MARK = "recovery"

# XLSX 表头
XLSX_HEADERS = ["title", "link"]

# 默认输出目录名（GUI 未选择保存路径时，回退到项目上级目录下的该子目录）
DEFAULT_OUTPUT_DIRNAME = "output"


# ======================== 数据结构 ========================
@dataclass
class CategoryItem:
    """一个待采集的分类：标题 + 列表页 URL。"""

    title: str
    url: str


@dataclass
class CollectConfig:
    """一次采集任务的全部输入。

    由 GUI 校验并组装完成后交给 CollectorRunner，采集端只读取本结构，
    不访问任何 tkinter 对象。
    """

    categories: List[CategoryItem] = field(default_factory=list)
    output_dir: str = ""
    # 0 表示不限制页数
    max_pages: int = 0
    # 空字符串表示走 HTML 正则模式，非空则走 XPath 模式
    xpath: str = ""
    # 浏览器参数（默认取本模块常量，GUI 可覆盖）
    proxy: Optional[str] = PROXY_SERVER
    port_range: tuple = PORT_RANGE
    disable_images: bool = DISABLE_IMAGES

    def summary(self) -> str:
        """返回任务摘要文本，供日志开头打印。"""
        lines = [
            f"共 {len(self.categories)} 个分类",
            f"最大页数: {self.max_pages if self.max_pages else '不限制'}",
        ]
        if self.xpath:
            lines.append(f"XPath: {self.xpath}")
        lines.append(f"保存目录: {self.output_dir}")
        lines.append("采集模式: 纯分页")
        return "\n".join(lines)
