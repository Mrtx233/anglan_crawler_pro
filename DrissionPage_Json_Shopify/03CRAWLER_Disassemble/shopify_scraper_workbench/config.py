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

# styles1 超过此长度会被 Excel 单元格上限（32767）截断，留出余量
MAX_STYLES_LENGTH = 32000

CSV_FIELDS = [
    "title", "name", "price1", "price2",
    "styles1", "styles2", "styles3",
    "src_links", "link-href", "details",
]

SIZE_VALUES = {
    "xs", "s", "m", "l", "xl", "xxl", "xxxl", "xxxxl",
    "2xl", "3xl", "4xl", "5xl",
    "xs/s", "s/m", "m/l", "l/xl",
    "one size", "os", "free size",
}

# ── styles 文本转义 ──────────────────────────────────────────
# styles 用 & # @ 作字段分隔符；选项值/选项名若含这些字符（如 "Tie & Square"）
# 会被误切成多个字段导致乱码。写入时对保留字符做反斜杠转义，解析时按层还原。
_STYLE_RESERVED = r"\&#@"

# 汇率缓存：由 network_utils.fetch_exchange_rates() 填充。
# 放在 config 是为了打破 network_utils <-> product_parser 的循环导入
# （product_parser 需要读汇率，network_utils 需要调 product_parser 的 format_price）。
_exchange_rates_cache = {}

PRICE_LIBRARY = [
    8.99, 9.01, 9.31, 9.33, 9.37, 9.63, 9.71, 9.72, 9.84, 9.86, 9.95, 9.99,
    10.03, 10.05, 10.11, 10.13, 10.15, 10.35, 10.36, 10.38, 10.54, 10.63, 10.64, 10.66,
    10.71, 10.78, 10.83, 10.94, 10.95, 10.99, 11.22, 11.36, 11.48, 11.65, 11.87, 11.95,
    11.99, 12.22, 12.25, 12.36, 12.41, 12.87, 12.92, 12.95, 12.99, 13.02, 13.13, 13.14,
    13.47, 13.69, 13.95, 13.99, 14.25, 14.32, 14.62, 14.83, 14.91, 14.95, 14.99, 15.31,
    15.52, 15.78, 15.95, 15.99, 16.34, 16.53, 16.66, 16.77, 16.95, 16.99, 17.11, 17.34,
    17.57, 17.95, 17.99, 18.62, 18.81, 18.84, 18.95, 18.99, 19.04, 19.25, 19.95, 19.99,
    20.14, 20.19, 20.34, 20.68, 20.74, 20.95, 20.99, 21.3, 21.47, 21.74, 21.95, 21.99,
    22.31, 22.34, 22.81, 22.95, 22.99, 23.14, 23.32, 23.95, 23.99, 24.22, 24.46, 24.63,
    24.95, 24.99, 25.37, 25.95, 25.99, 26.18, 26.34, 26.95, 26.99, 27.31, 27.54, 27.95,
    27.99, 28.11, 28.95, 28.99, 29, 29.31, 29.95, 29.99, 30.54, 30.95, 30.99, 31.26,
    31.52, 31.95, 31.99, 32.14, 32.24, 32.95, 32.99, 33.62, 33.84, 33.95, 33.99, 34.18,
    34.62, 34.95, 34.99, 35, 35.14, 35.33, 35.95, 35.99, 36.47, 36.95, 36.99, 37.21,
    37.95, 37.99, 38.14, 38.52, 38.95, 38.99, 39, 39.54, 39.95, 39.99, 40.02, 40.15,
    40.95, 40.99, 41.32, 41.95, 41.99, 42.51, 42.95, 42.99, 43.16, 43.95, 43.99, 44.82,
    44.95, 44.99, 45.17, 45.95, 45.99, 46.57, 47.82, 48.36, 49, 49.31, 50.03, 50.13,
    50.81, 55, 59, 69, 79, 85, 89, 99, 109, 119, 129, 189, 219, 259, 299, 319,
    349, 369, 399, 459, 499, 519, 599,
]

SHOPIFY_COLUMNS = [
    "Link-Href", "Handle", "Title", "Body (HTML)", "Collection",
    "Vendor", "Type", "Tags", "Published", "Option1 Name",
    "Option1 Value", "Option2 Name", "Option2 Value", "Option3 Name",
    "Option3 Value", "Variant SKU", "Variant Grams",
    "Variant Inventory Tracker", "Variant Inventory Qty",
    "Variant Inventory Policy", "Variant Fulfillment Service",
    "Variant Price", "Variant Compare At Price",
    "Variant Requires Shipping", "Variant Taxable",
    "Variant Barcode", "Image Src", "Image Position",
    "Image Alt Text", "Gift Card", "SEO Title",
    "SEO Description", "Google Shopping - Google Product Category",
    "Google Shopping - Gender", "Google Shopping - Age Group",
    "Google Shopping - MPN", "Google Shopping - AdWords Grouping",
    "Google Shopping - AdWords Labels", "Google Shopping - Condition",
    "Google Shopping - Custom Product", "Google Shopping - Custom Label 0",
    "Google Shopping - Custom Label 1", "Google Shopping - Custom Label 2",
    "Google Shopping - Custom Label 3", "Google Shopping - Custom Label 4",
    "Variant Image", "Variant Weight Unit", "Variant Tax Code",
    "Cost per item",
]

WP_HEADERS = [
    "Type", "SKU", "Name", "Published", "Visibility in catalog", "Description", "In stock?", "Stock",
    "Sale price", "Regular price", "Categories", "Tags", "Images", "Parent", "Position",
    "Attribute 1 name", "Attribute 1 value(s)", "Attribute 1 visible", "Attribute 1 global",
    "Attribute 2 name", "Attribute 2 value(s)", "Attribute 2 visible", "Attribute 2 global",
    "Attribute 3 name", "Attribute 3 value(s)", "Attribute 3 visible", "Attribute 3 global",
]

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
# 输入 XLSX 所在目录名
INPUT_DIRNAME = "01INPUT_XLSX"

# 输出 XLSX 所在目录名（build_output_file_path 会把输入路径中的
# INPUT_DIRNAME 整段替换为本值）
OUTPUT_DIRNAME = "02OUTPUT_XLSX"

