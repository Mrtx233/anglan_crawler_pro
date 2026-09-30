from __future__ import annotations

LINK_MAX_RETRIES = 5


RETRY_DELAY = 5


PAGE_LOAD_TIMEOUT = 30


DELAY = 0.5


PRODUCT_MAX_RETRIES = 5


RETRY_BASE_DELAY = 30


SAVE_EVERY = 20


SKIP_OPTIONS = None


# styles1 超过此长度会被 Excel 单元格上限（32767）截断，留出余量
MAX_STYLES_LENGTH = 32000


# 合并时剔除 name 包含这些关键词的商品，按整体字面量匹配，大小写不敏感
EXCLUDE_NAME_KEYWORDS = ["Gift Card"]


INPUT_DIR_NAME = "01INPUT_XLSX"


OUTPUT_DIR_NAME = "02OUTPUT_XLSX"

PROXY_URL = "http://127.0.0.1:7897"
