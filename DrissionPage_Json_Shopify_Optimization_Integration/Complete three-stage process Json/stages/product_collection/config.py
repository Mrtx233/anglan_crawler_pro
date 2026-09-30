# ======================== 配置层 ========================
#
# 职责: 集中管理全部常量与全局可变配置。
# 依赖: 无（依赖图的叶子模块，禁止导入任何其他业务模块）。
#
# 全局可变配置说明:
#   KEEP_POSITIONS / SKIP_POSITIONS 需要按「每个输入文件」临时切换。
#   原实现用 globals()[...] 在同一模块内改写；拆分后跨模块 globals()
#   已不可用，改为由 product_parser 提供 with_image_filter() 上下文管理器，
#   内部读写本模块的这两个属性，语义与原实现完全一致。
# ========================================================

# 本模块不依赖任何第三方库与业务模块

# ════════════════════════════════════════════════════════════
# 全局采集配置
# ════════════════════════════════════════════════════════════
DELAY = 0.5

MAX_RETRIES = 5

RETRY_BASE_DELAY = 30

SAVE_EVERY = 20

SKIP_POSITIONS = None

KEEP_POSITIONS = None

SKIP_OPTIONS = ["ships from"]  # 屏蔽"发货地"选项，它不属于商品变体

# 选项值里含这些关键字（子串、不分大小写）时直接剔除。
# "Custom Size" 是定制尺码，不是可履约的真实变体，导出到 Shopify 会变成
# 一个卖不了的 SKU；按子串匹配可一并覆盖 "Custom Made"/"Custom Order" 等写法。
# 过滤在 product_record 产出记录时生效，见 _filter_option_values。
EXCLUDED_VALUE_KEYWORDS = ["custom"]

# 分类级图片过滤的「显式不过滤」记号。填在分类行的某一栏表示该栏不过滤，
# 同时让该分类行进入独立模式（不再继承文件行设置）。
# 解析逻辑见 scraper_runner.ScraperRunner._resolve_image_filter。
IMAGE_FILTER_NONE = "-"

# 不再使用：第二阶段已改为 JSON 输出，没有 Excel 单元格（32767）长度限制。
# 保留常量仅为兼容旧调用与便于回退到 XLSX 输出。
MAX_STYLES_LENGTH = 32000

# 旧 XLSX 中间表的列（parse_product 仍在产出这一形状，供回退与第三阶段读取）
CSV_FIELDS = [
    "title", "name", "price1", "price2",
    "styles1", "styles2", "styles3",
    "src_links", "link-href", "details",
]

# 第二阶段 JSON 记录的字段（product_record.parse_product_record 的产出）
RECORD_FIELDS = [
    "category", "name", "link_href", "details", "options",
    "default_price", "variant_images", "variant_overrides", "src_links",
]

SIZE_VALUES = {
    "xs", "s", "m", "l", "xl", "xxl", "xxxl", "xxxxl",
    "2xl", "3xl", "4xl", "5xl",
    "xs/s", "s/m", "m/l", "l/xl",
    "one size", "os", "free size",
}

# ── styles 文本转义 ──────────────────────────────────────────
# styles 用 & # @ 作字段分隔符；选项值/选项名若含这些字符（如 "Tie & Square"）
# 会被误切成多个字段导致乱码。写入时对保留字符做反斜杠转义。
_STYLE_RESERVED = r"\&#@"

# 汇率缓存：由 network_utils.fetch_exchange_rates() 填充。
# 放在 config 是为了打破 network_utils <-> product_parser 的循环导入
# （product_parser 需要读汇率，network_utils 需要读同一份缓存以折算汇率）。
_exchange_rates_cache = {}

UI_COLORS = {
    "background": "#F4F7FB",
    "card": "#FFFFFF",
    "card_alt": "#F8FAFC",
    "border": "#E2E8F0",
    "border_focus": "#93C5FD",
    "text": "#0F172A",
    "text_secondary": "#475569",
    "text_muted": "#94A3B8",
    "primary": "#2563EB",
    "primary_hover": "#1D4ED8",
    "primary_light": "#EFF6FF",
    "primary_border": "#BFDBFE",
    "success": "#16A34A",
    "success_light": "#F0FDF4",
    "success_border": "#BBF7D0",
    "warning": "#D97706",
    "warning_light": "#FFFBEB",
    "warning_border": "#FDE68A",
    "danger": "#DC2626",
    "danger_hover": "#B91C1C",
    "danger_light": "#FEF2F2",
    "danger_border": "#FECACA",
    "disabled": "#CBD5E1",
    "log_background": "#0F172A",
    "log_panel": "#111827",
    "log_text": "#D1D5DB",
}

UI_FONT = "Microsoft YaHei UI"

MONO_FONT = "Consolas"

# ======================== 输入/输出目录约定 ========================
# 输入目录名（第一阶段的 01INPUT_XLSX 及第二阶段的来源目录）
INPUT_DIRNAME = "01INPUT_XLSX"

# 输出 XLSX 所在目录名（build_output_file_path 会把输入路径中的
# INPUT_DIRNAME 整段替换为本值）。中间表放在任务目录的同级而非任务目录内，
# 避免与第一阶段的链接 JSON 混在一起。
OUTPUT_DIRNAME = "02OUTPUT_XLSX"
