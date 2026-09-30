# anglan_crawler_pro — Shopify 商品采集与数据转换工具集

面向跨境电商运营的**商品数据采集流水线**：从站点链接库管理，到列表页链接采集、商品数据抓取、变体解析编码，最终产出可直接导入 **Shopify** 与 **WooCommerce (WP)** 的商品文件。

所有路线共享同一套 10 字段数据契约（见[下节](#10-字段数据契约)），因此上游采集器可以互换，下游转换逻辑可以复用。

| 路线 | 技术栈 | 适用场景 | 当前状态 |
|---|---|---|---|
| **GUI 工作流** | DrissionPage + Tkinter | 标准 Shopify 站点的日常批量操作 | 单文件版仅存 `Shopify_JSON_GUI.py`、`detail_link_GUI.py`；加固版与分页版已归档至 `03CRAWLER/clean/`，活跃维护版在 `03CRAWLER_Disassemble/` |
| **Scrapy 工作流** | Scrapy 2.17 | 定制化爬虫开发，模块化共享 utils | 5 个 spider（3 个「域名+分类」目录） |
| **非标准站点路线** | DrissionPage + 站点专属解析 | Headless / Next.js / Magento / WooCommerce / Shopline 等无 `.json` 接口站点 | 5 个已接入站点脚本 |
| **styles 单脚本线** | 纯 pandas / openpyxl，无 GUI | 对已有 styles 表做原价修正 + 递增/分档价格匹配 + 转换 | `Data logic/`（实验与基准），`standard_collection_process_Preposition`（同一思路的分档实现） |

> 本文核对日期 **2026-09-29**，所有行数、常量值、目录与结论均以当前工作区代码为准。与 2026-09-17 版本的差异见[修订记录](#本次修订要点)。

---

## 目录

- [环境要求](#环境要求)
- [项目结构](#项目结构)
- [数据流总览](#数据流总览)
- [10 字段数据契约](#10-字段数据契约)
- [Styles 编码格式](#styles-编码格式)
- [四阶段转换流水线](#四阶段转换流水线)
- [价格匹配：两套并存的算法](#价格匹配两套并存的算法)
- [工具说明](#工具说明)
  - [链接库管理](#链接库管理)
  - [链接采集](#链接采集)
  - [商品采集](#商品采集)
  - [非 Shopify 站点采集](#非-shopify-站点采集)
  - [styles 单脚本线（Data logic）](#styles-单脚本线data-logic)
  - [转换与辅助工具](#转换与辅助工具)
- [模块化拆分版](#模块化拆分版)
- [运行方式与测试](#运行方式与测试)
- [输入输出目录组织](#输入输出目录组织)
- [文档索引](#文档索引)
- [已知问题与修复状态](#已知问题与修复状态)
- [本次修订要点](#本次修订要点)

---

## 环境要求

- **Python 3.12**（`.venv/` 已就绪；`standard_collection_process{,_Preposition}/requirements.txt` 声明最低 3.11）
- **本机 Chromium/Chrome** + **tkinter**（GUI 工具）
- **本地代理**：`http://127.0.0.1:7897`（Clash 等）。代理地址写在各工具的 `PROXY_SERVER` / `PROXY_URL` 常量中，设为 `None` 或 `""` 表示直连

```bash
pip install -r requirements.txt
```

| 依赖 | 版本 | 用途 |
|---|---|---|
| `scrapy` | 2.17.0 | 商品详情页 JSON 采集框架 |
| `DrissionPage` | 4.1.1.4 | 浏览器自动化（链接采集、JSON 抓取、非标准站点渲染） |
| `lxml` | 6.1.1 | HTML 解析 |
| `pandas` | 2.3.3 | 表格读写与格式转换 |
| `openpyxl` | 3.1.5 | Excel XLSX 读写 |
| `itemadapter` | 0.13.1 | Scrapy Item 与 pipeline 适配 |

三点与依赖有关的实际情况：

1. **根 `requirements.txt` 没有 `pytest`，`.venv` 中也没有**。`standard_collection_process` 的测试是 `unittest` 风格，用 `python -m unittest` 即可运行，无需安装 pytest（见[运行方式与测试](#运行方式与测试)）。
2. `Utils/Single function/图片转JPG.py` 会在运行时按需 `pip install Pillow`，并尝试加载可选的 `pillow-heif`；缺 `pillow-heif` 时 `.heic/.heif` 跳过，其余格式不受影响。
3. `standard_collection_process{,_Preposition}/requirements.txt` 单独锁版本，其中 **pandas 写的是 3.0.3**，与根目录的 2.3.3 不一致，按各自目录的 requirements 安装。

---

## 项目结构

```
anglan_crawler_pro/
├── requirements.txt
├── README.md
│
├── Assembly line/                        # ① 链接库管理（流水线最前置环节）
│   ├── jsonl_gui.py                      #    链接库 GUI（2745 行，侧边栏导航 + 多页面）
│   └── links.jsonl                       #    链接库数据（1979 条，JSONL，9 字段）
│
├── DrissionPage_Json_Shopify/            # ② GUI 工作流主目录
│   ├── 00Historical data/                #    归档的历史批次数据（01INPUT_XLSX / 02OUTPUT_XLSX）
│   ├── 01INPUT_XLSX/                     #    输入：批次 A0730_2 A0801 A0812 A0901 A0910 A0916
│   ├── 02OUTPUT_XLSX/                    #    输出：批次 A0801 A0812 A0901 A0910 A0916
│   │
│   ├── 03CRAWLER/                        #    单文件版工具
│   │   ├── Shopify_JSON_GUI.py           #      2463 行  商品采集（基线版，当前在 03CRAWLER 根目录）
│   │   ├── detail_link_GUI.py            #      1193 行  链接采集（无限滚动站点）
│   │   ├── styles数据处理流程报告.md      #      1107 行  四阶段流水线详解（GUI/旧算法口径）
│   │   └── clean/                        #      归档快照：a703afb 纯移动（0 行改动），无任何代码引用
│   │       ├── link_collector_pagination.py    #   934 行
│   │       ├── shopify_scraper_workbench.py    #  2587 行
│   │       └── standard_collection_process.py  #  3206 行（链接采集+商品采集+转换一体版）
│   │
│   ├── 03CRAWLER_Disassemble/            #    模块化拆分版（界面与业务分离，活跃维护）
│   │   ├── link_collector_pagination/    #       6 模块 + README(258)，已实现原子写入与 recovery
│   │   ├── shopify_scraper_workbench/    #       9 模块 + README(350)，新增 scraper_runner.py
│   │   ├── standard_collection_process/  #       54 文件 + 55 项测试（唯一带测试的版本）
│   │   └── standard_collection_process_Preposition/  #  同源副本 + 域名替换/关键词剔除/百分位分档价格
│   │
│   ├── 04非 Shopify 标准站点采集/
│   │   ├── non_shopify_utils.py          #      187 行  公共工具（价格/汇率/图片/HTML）
│   │   ├── 非Shopify站点采集脚本开发流程.md  #   1068 行  新站点接入方法论
│   │   └── 非 Shopify 站点/
│   │       ├── snitch_crawler_GUI.py         #   389 行  Shopify + XPath 渲染页
│   │       ├── dobell_crawler_GUI.py         #   503 行  Magento PWA
│   │       ├── joyfit_crawler_GUI.py         #   502 行  Shopline（SSR HTML）
│   │       ├── teeshoppen_crawler_GUI.py     #   534 行  Shopify Hydrogen
│   │       └── hespokestyle_crawler_GUI.py   #   546 行  WordPress + WooCommerce
│   │
│   └── Data logic/                       # ③ styles 单脚本线（原价修正 + 递增价格匹配）
│       ├── Styles_递增价格匹配.py        #      700 行  无 GUI 无参数，改 INPUT_FILE 直接跑
│       ├── Data logic.md                 #      226 行  该脚本的逐步说明（含 2 处过期描述，见文档索引）
│       └── styles/                       #      lovedandco 的一次运行样例产物（原价/ 与 价格匹配/）
│
├── ShopifyJsonDetailPage/                # ④ Scrapy 工作流
│   ├── scrapy.cfg
│   └── ShopifyJsonDetailPage/
│       ├── items.py                      #    ShopifyJsonItem 定义
│       ├── pipelines.py                  #    XLSX 输出 + 断点续采
│       ├── middlewares.py
│       ├── settings.py                   #    配置 + 自定义 Spider Loader
│       ├── output_details/               #    商品 JSON 详情输出
│       ├── output_styles/                #    Styles 编码输出（按 spider/时间戳分目录）
│       ├── spiders/                      #    按「域名+分类」中文子目录组织
│       │   ├── kestrs.com男装/                   mnml_la.py
│       │   ├── midsummera.com男女装日常休闲款/     chatsworthboutique / folkclothing / harperandco
│       │   └── vestf.com职业装-女/                mmlafleur.py
│       └── utils/
│           ├── json_variants.py          #    Shopify JSON → Styles 编码（411 行）
│           ├── link_collector.py         #    DrissionPage 链接采集（230 行）
│           ├── shopify_converter.py      #    Styles → Shopify 格式 + 价格匹配（382 行）
│           ├── spider_helpers.py         #    通用辅助函数
│           └── spider_loader.py          #    支持中文目录名的 Spider Loader
│
└── Utils/                                # ⑤ 独立工具
    ├── switch_gui.py                     #   1505 行  8 合 1 转换工具 GUI
    └── Single function/                  #   b439ea5 起统一中文命名的单功能脚本
        ├── 价格修正.py                    #    392 行  Sale==Regular 时 Regular×1.2（GUI + 命令行）
        ├── 图片转JPG.py                   #    290 行  批量图片格式转换，联动改写 CSV 后缀
        ├── 按标题拆分Excel.py              #    369 行  按 title 分组、组内价格升序轮询拆分
        ├── 按SKU抽样.py                   #     70 行  从 wp-*.csv 抽 N 个主体商品及其变体
        ├── 描述字段清洗.py                 #     52 行  清理 Description 中的编号与名人推荐段
        └── 提取本地图片.py                 #     73 行  按 styles1 引用精简 src_links
```

> 说明：`03CRAWLER_Disassemble/*.md` 那四份逐文件源码归档已在 **b439ea5**（2026-09-21）删除，源码不再需要与归档同步。

---

## 数据流总览

```
① 链接库 links.jsonl ──── Assembly line/jsonl_gui.py 录入、审核、标注
        │  （id/batch/type/link/status/created_at/note/updated_at/approver）
        ▼
   列表页 URL（status 标为「可使用」的站点）
        │
        ▼
② 链接采集 ──── detail_link_GUI.py（无限滚动） / link_collector_pagination（?page=N 分页）
        │         产物：{domain}.xlsx（title, link）+ links_index.json（分类 URL 索引）
        ▼
   01INPUT_XLSX/{批次}/{操作员}/{域名+分类}/*.xlsx
        │
        ▼
③ 商品采集 ──── 03CRAWLER_Disassemble/shopify_scraper_workbench（标准 Shopify）
        │         或 04非 Shopify 标准站点采集/（非标准站点）
        │         或 Scrapy spiders
        │         请求 {product}.json → 解析变体/价格/图片/描述
        ▼
   Styles XLSX（10 字段契约，自定义编码承载变体）
        │
        ▼
④ 「合并并转换」── 人工断点，操作员先检查各来源站拆分表
        │
        ├──► Shopify 导入表（49 列，含 Handle/SKU/库存/Google Shopping）
        │         │
        │         ▼
        │    价格匹配（两套算法，见专节）
        │         │       GUI 线产物：原价 / 价格匹配 / 未匹配到的价格（三态，xlsx + csv）
        │         ▼
        └──► WooCommerce (WP) CSV（workbench 线 27 列；两个拆分包 26 列）
```

---

## 10 字段数据契约

**全项目统一契约**，标准 Shopify 管线与非标准站点管线输出完全一致，因此下游转换逻辑可直接复用：

```python
CSV_FIELDS = [
    "title",      # 输入 xlsx 第一列原样透传（分类名）
    "name",       # 商品真实名称（非空、非 Skeleton、非推荐商品）
    "price1",     # 实际售价，统一转为 USD
    "price2",     # 真实存在的划线价；不存在则 price2 = price1（禁止人为制造）
    "styles1",    # 变体编码（无变体时为空串）
    "styles2",    # 保留，当前流程恒为 ""
    "styles3",    # 保留，当前流程恒为 ""
    "src_links",  # 商品主图，img1#img2#img3
    "link-href",  # 商品详情页 URL（断点续采依据）
    "details",    # 商品详情 HTML（已清洗）
]
```

实测各实现中的列定义常量：`CSV_FIELDS` 10 列、`SHOPIFY_COLUMNS` 49 列（全实现一致）、`WP_HEADERS` **27 列**（`shopify_scraper_workbench` 及其归档的 `03CRAWLER/clean/`、`Shopify_JSON_GUI.py`，含 `Tags`）或 **26 列**（`standard_collection_process` 与 `..._Preposition` 的 `resources/columns.py`，去掉 `Tags`）。`PRICE_LIBRARY` 为 215 个价位（8.99–599），四份副本内容一致。

多维 Option 统一放入 `styles1` 的 segment 中表达，不拆分到 styles2/styles3。

---

## Styles 编码格式

连接采集层与转换层的中间表示，用分隔符承载「选项名 / 选项值 / 子选项 / 双价格 / 图片」：

```
OptionName#Value1&SubOption&SubValue$price1$price2@imageURL#Value2&...
```

示例：

```
Color#Red&Size&S$29.99$35.99@https://cdn.shopify.com/...jpg#Blue&Size&M$29.99$35.99@...
```

| 分隔符 | 含义 |
|---|---|
| `#` | 选项值之间 |
| `&` | 选项名与子选项之间 |
| `$` | 价格字段 |
| `@` | 图片字段 |

**特殊字符转义**：选项名或选项值本身含 `\ & # @` 时（如 `Tie & Square`）会用反斜杠转义，保留字符集为 `_STYLE_RESERVED = r"\&#@"`（各实现一致，`Utils/switch_gui.py` 的 WP 还原 styles 也复用同一套转义）。

**尺码统一规则**：`_is_size_option` + `SIZE_VALUES` 白名单（`xs/s/m/l/xl/xxl/xxxl/xxxxl`、`2xl`–`5xl`、`xs/s` 等区间值、`one size`/`os`/`free size`）识别尺码维度。

**Type → Style 改名**：`Type`/`Option Type` 类选项统一重命名为 `Style`；`Ships From`（发货地）通过 `SKIP_OPTIONS = ["ships from"]` 屏蔽，不进入变体笛卡尔积。注意两个拆分包的 `config.py` 里 `SKIP_OPTIONS = None`（不屏蔽），单文件版仍为列表——这条配置在不同实现间已经分叉。

**长度上限**：`styles1` 是单个 Excel 单元格，超长会被 32767 上限**静默截断**、直接损坏数据。变体组合数极多的商品会撞上这一限制，因此采集侧加了 `MAX_STYLES_LENGTH = 32000`：编码超过该长度即抛 `InvalidProductError`，整个商品跳过，不写入来源表和合并表，日志给出实际长度与上限。取 32000 而非 32767 是为边界问题和 `styles2`/`styles3` 可能启用留余量。

> ⚠️ **状态未变**：这道检查只存在于 `03CRAWLER_Disassemble/` 的三个包（`standard_collection_process{,_Preposition}` 的 `config.py` + `processing/product.py`，`shopify_scraper_workbench` 的 `config.py` + `product_parser.py`）。`03CRAWLER/Shopify_JSON_GUI.py` 与归档的 `03CRAWLER/clean/*.py` 全仓 grep 无 `MAX_STYLES_LENGTH`，仍存在超长 `styles1` 被截断的风险。

---

## 四阶段转换流水线

详见 [`03CRAWLER/styles数据处理流程报告.md`](DrissionPage_Json_Shopify/03CRAWLER/styles数据处理流程报告.md)（1107 行，含逐行处理流程、SKU 生成详解、完整数据示例；**按 GUI/旧价格算法口径撰写**）。

| 阶段 | 入口函数 | 产物 |
|---|---|---|
| **一 采集** | `parse_product()` | 10 字段 xlsx（按来源站拆分） |
| **断点** | 「合并并转换」按钮 | `{分类}_合并.xlsx` |
| **二 Shopify 转换** | `_styles_to_shopify()` / `shopify.py` | 49 列 Shopify 格式（`SHOPIFY_COLUMNS`） |
| **三 价格匹配** | `_price_match()` / `match_prices()` / 百分位分档 | 原价 / 价格匹配 / 未匹配到的价格 |
| **四 WP 转换** | `_shopify_to_wp()` / `woocommerce.py` | WooCommerce 导入 CSV（26 或 27 列） |

**「合并并转换」是刻意设计的人工断点**：采集完成后不自动进入阶段二，而是把结果存入 `pending_merge` 并启用按钮，等操作员检查各来源站拆分表、发现漏采错采可重跑采集，避免坏数据流入后续转换。

---

## 价格匹配：两套并存的算法

代码里目前并存**两条价格匹配线**，规则不同、作用对象不同、产物目录也不同。选错会直接导致产物对不上，先确认自己在哪条线上。

### A. Shopify 表匹配（GUI 主线，历史实现）

作用在**阶段二产出的 49 列 Shopify 表**上，改写 `Variant Price` / `Variant Compare At Price`。存在四处，逻辑一致：

| 位置 | 函数 |
|---|---|
| `03CRAWLER/Shopify_JSON_GUI.py` | `_price_match()` |
| `03CRAWLER/clean/standard_collection_process.py` | `_price_match()` |
| `03CRAWLER_Disassemble/shopify_scraper_workbench/converter.py` | `_price_match()` |
| `03CRAWLER_Disassemble/standard_collection_process/processing/pricing.py` | `match_prices()`（同一算法的模块化版，52 行） |

规则：

- 原价窗口 = `original × [0.75, 1.25]`（对称 ±25%）
- 命中方式：取窗口内**最小候选**（`next(...)` / `candidates[0]`），**不是**「最近值」
- 按 `groupby("Handle")` 分组，组内按原价升序遍历
- `global_used_prices` **全文件硬去重**：一个价位在整个文件里只用一次
- 划线价 = `round(匹配价 × 1.2, 2)`，并新增 `匹配价格` 列
- 产物三态：`_Shopify_原价` / `_Shopify_价格匹配` / `_Shopify_未匹配到的价格`，**xlsx + csv 双份**，分别落在 `xlsx/` 与 `csv/` 子目录

### B. styles 表匹配（Data logic 线，较新）

作用在**阶段一产出的 10 列 styles 表**上，直接改写 `price1` / `price2` 以及 `styles1` 段内嵌价，在 Shopify 转换**之前**完成匹配。两个实现同源：**`Data logic/Styles_递增价格匹配.py` 是基准蓝本**，`standard_collection_process_Preposition/processing/pricing.py`（808 行）在其上换成了百分位分档。

**B-1 `Data logic/Styles_递增价格匹配.py`**

- 原价 = `price2`（D 列），为空或 ≤0 时回退 `price1`
- 窗口 = `original × [0.6, 1.1]`（`LOWER_RATIO=0.6`、`UPPER_RATIO=1.1`，**非对称**）
- `original > 599`（`MAX_LIB_PRICE`）→ 全部变体封顶取 `599`
- 候选数 ≥ 变体数：取**最小的 V 个**候选；候选不足：按比例重复 `candidates[i*n//V]`
- **不做跨商品去重**（`used_prices` 只用于统计哪些价位没用过，不从候选中剔除），因此同一价格可被多商品复用
- 行级聚合：`price1 = min(本行分配价)`、`price2 = max(本行分配价)`；段内 price2 统一写本行最大值
- 产物目录：`原价/`、`价格匹配/`（**Shopify 仅 CSV**，WP 为同目录 `wp-<文件名>.csv`）+ `价格匹配/<stem>_未匹配到的价格.csv`（单列）
- 「**递增**」的含义：同一商品各变体按候选**升序依次分配**，变体价单调不降

**B-2 `standard_collection_process_Preposition/processing/pricing.py`（百分位分档，b5804fd）**

- `fix_original_prices()` 先做原价修正（行级 `price2 < price1` 拉平、`styles1` 段内嵌价拉平）
- `_get_price_percentiles()`：按**全文件唯一原价**排序取 rank（最低 0、最高 1，重复价同档，只有一个原价时统一 0.5）
- `PRICE_RATIO_RULES` 五档窗口（原价越靠低价分位，允许的折扣越深）：

  | 百分位 | 窗口 |
  |---|---|
  | ≤ 0.10 | `[0.20, 0.80]` |
  | ≤ 0.30 | `[0.40, 0.90]` |
  | ≤ 0.70 | `[0.60, 1.10]` |
  | ≤ 0.90 | `[0.70, 1.30]` |
  | ≤ 1.00 | `[0.80, 1.50]` |

- `_get_target_price()`：目标价 = 窗口下沿 + 窗口宽度 × 百分位
- `_get_candidate_pool()`：三层回退——区间内库价 → 两侧各扩张 `FALLBACK_EXPANSIONS = (0.10, 0.20)` → 取离目标价最近的 `max(12, 变体数 × 4)` 个
- `_assign_variant_prices()`：按 `(全局使用次数, 归一化距离, 价格)` 评分选择，**usage 计数而非硬去重**，候选不足时均衡重复，最后 `sorted(selected)` 升序回给各变体（保留「递增」语义）
- 行级 `price1 = min`、`price2 = max`；**没有 ×1.2 划线价**，划线价即本行最大匹配价
- 输出与 B-1 相同的 `原价/`、`价格匹配/` 结构
- 已知副作用：**结果与行序相关**（该包 README 已承认）

演进链（便于日后回看）：`7afc939` 建包时窗口 `[0.5, 1.25]` → `477b3a4` 把 Preposition 的 `LOWER_RATIO` 收紧为 `0.75` → `b5804fd` 换成百分位分档（未匹配价位数 144 → 108）。

### 差异速查

| 维度 | A：Shopify 表匹配 | B-1：Data logic 脚本 | B-2：Preposition 分档 |
|---|---|---|---|
| 作用对象 | Shopify 49 列 `Variant Price` | styles 10 列 `price1/price2/styles1` | styles 10 列 |
| 窗口 | `×[0.75, 1.25]` 对称 | `×[0.6, 1.1]` | 五档 `0.20–1.50` 按百分位 |
| 取价 | 窗口内最小候选 | 窗口内升序前 V 个 | 离目标价最近 |
| 跨商品去重 | 全局硬去重（每价一次） | 无 | usage 计分，可复用 |
| 划线价 | 匹配价 × 1.2 | 本行最大匹配价 | 本行最大匹配价 |
| 产物 | `xlsx/`+`csv/` 三态双格式 | `原价/`+`价格匹配/`，仅 CSV | 同 B-1 |
| 确定性 | 依赖 random + 行序 | 依赖 random + 行序 | 依赖行序 |

三条线都含 `random` 生成的 Handle/SKU，**重复运行同输入结果不同**，不能靠重跑做增量。

---

## 工具说明

### 链接库管理

#### `Assembly line/jsonl_gui.py`（2745 行）

采集前的站点链接池管理与审核工具，`LinkBrowserApp` 类含 44 个方法，**侧边栏导航 + 多页面切换**架构（非选项卡）。

**数据格式**（`links.jsonl`，每行一条 JSON，9 个字段；当前 **1979 条**，批次分布 A0819 1521 / A0820 360 / A0910 98）：

```json
{
  "id": "00001",
  "batch": "A0819",
  "type": "运动服装",
  "link": "https://www.silverbackgymwear.com",
  "status": "历史，不在启用",
  "created_at": "2026-08-19 11:04:02",
  "note": "",
  "updated_at": "2026-08-20 16:00:00",
  "approver": ""
}
```

字段顺序有两种写法（1851 条为 `...note, updated_at, approver`，128 条为 `...updated_at, note, approver`），读取按 key 取值不受影响。

**状态取值**：`未使用`（默认）/ 预设 `可使用`、`待定`、`量少`、`logo`、`404` / 自定义状态（当前实际出现的还有 `历史，不在启用`、`审核中`、`待审核`、`不可用`、`印度`、`严格复审`）。每种状态有独立配色（`STATUS_THEME`）。

**页面构成**：

- **链接库页**（`_build_library_page`）：`批次 → 分类 → 链接` 三级下钻
  - 表格（`_build_table`）+ 详情面板（`_build_detail`）+ 搜索框 + 状态筛选下拉
  - 批量操作：状态快捷标注、自定义状态、备注、**审核人（approver）**，支持多选批量修改，即时持久化
  - 行内「复制」按钮一键复制链接（含 `_toast` 气泡提示）；双击 `on_double_click` 打开浏览器
  - 列宽随窗口自适应（`_resize_columns`），状态着色显示
- **录入页**（`_build_add_page`）：粘贴多行「链接 分类」批量录入
  - `normalize_link` 域名归一化去重（去协议/www/末尾斜杠）
  - `parse_input_line` 解析行，自动 id 递增，状态默认「未使用」

### 链接采集

两条实现对应 Shopify 列表页的两种加载模式，**互不回退**。

#### `detail_link_GUI.py` — 无限滚动站点（1193 行）

针对瀑布流/无限滚动加载的列表页，仍在 `03CRAWLER/` 根目录。

- **页面就绪判断**：`scroll_to_bottom` 滚动到底 + 等待 5 秒 + 连续 3 次高度不变则停止
- **无限滚动专项处理**：`is_infinite_scroll_page` 探测 → `collect_infinite_scroll_links` 采集；含 `_click_load_more_in_product_list`（点「加载更多」）、`_scroll_product_list_to_end`（滚动内部容器）、`_get_scroll_metrics`（读取滚动度量）
- **两种提取模式**：
  - **Handle 模式**（XPath 留空）：`extract_product_links` 合并三种来源——`var meta.products` JSON 解析 + `"handle":"xxx"` 正则 + DOM 链接，去重后校验 `/products/{handle}` 确实存在于 HTML
  - **XPath 模式**（用户填写 XPath）：`extract_links_by_xpath` 用 `document.evaluate` 执行，支持属性节点与元素节点
- **容错**：`_run_js_with_retry` 对 `ContextLostError`（页面刷新导致上下文丢失）自动重试
- **断点续采**：`load_existing_domain_rows` 读取已有输出合并去重，不覆盖；`update_link_index` 维护 `links_index.json`

配置常量：`MAX_RETRIES=3`、`RETRY_DELAY=5`、`PAGE_LOAD_WAIT=1.5`、`INFINITE_SCROLL_WAIT=1.2`、`INFINITE_SCROLL_STABLE_ROUNDS=5`、`INFINITE_SCROLL_MAX_ROUNDS=80`。

#### `link_collector_pagination` — `?page=N` 分页站点

同一功能的三种实现并存，改动请认准目录：

| 版本 | 位置 | 说明 |
|---|---|---|
| 单文件归档 | `03CRAWLER/clean/link_collector_pagination.py`（934 行） | a703afb 移入归档，此后无提交触及 |
| 模块化（活跃） | `03CRAWLER_Disassemble/link_collector_pagination/`（6 模块，共 1760 行） | 含关窗等待、原子写入、recovery、保存报告 |
| Scrapy 侧 | `ShopifyJsonDetailPage/.../utils/link_collector.py`（230 行） | Spider 内调用 |

- **翻页采集**：`build_page_url` 用 `parse_qsl`/`urlencode`/`urlunparse` 安全改写查询串，逐页 `?page=1,2,3...` 抓取，最大页数可配置（留空=不限制；GUI 不接受 0/负数/非整数）
- **三级正则提取**：`extract_product_links` 依次尝试 meta JSON → handle 正则 → DOM 链接，按出现顺序去重；XPath 非空时**只用** XPath
- **性能优化**：浏览器启动带 `--blink-settings=imagesEnabled=false`，只采集 HTML/JSON/链接
- **关窗等待保存**（`main.py`，9c4e94e）：`_on_close` 先 `interrupt`，再用 `root.after` 轮询等待采集线程结束，最后销毁窗口；`CLOSE_WAIT_TIMEOUT = 60` 秒超时兜底后强制退出并提示
- **保存报告**：`CollectorRunner` 保存后经 `on_save_report` 回调把 `SaveReport` 交回 GUI（`main.py:_on_save_report`），显示保存摘要；失败或产生 recovery 时弹窗提示
- **原子写入与 recovery**（`file_utils.py`）：XLSX 与 JSON 索引都先写同目录临时文件再 `os.replace`（`_atomic_replace`、`_atomic_write_json`）；JSON 读取失败先备份为 recovery；XLSX 写入失败也另存 recovery（`build_recovery_file_path`）；域名文件名做 Windows 合法化清洗
- 分层约束：`collector.py` **不导入 tkinter**，业务可脱离 GUI 调用

#### 公共输出

按域名分组，每个域名一个 `{domain}.xlsx`（列：`title`, `link`），同目录写 `links_index.json` 记录「分类名 → 列表页 URL」映射。分类名可对应多个 URL（同域名下重名分类，如 `slessic.com` 的多个 `Classic` 列表页），值一律为数组，按提交顺序排列并去重；旧版单字符串写法由 `_normalize_link_index` 在下次写入时自动升级。

### 商品采集

#### `Shopify_JSON_GUI.py` — 基线版（2463 行，`03CRAWLER/` 根目录）

读取 `01INPUT_XLSX/` 的产品链接 Excel，逐条请求 `.json` 接口，解析 Shopify 产品数据写入 `02OUTPUT_XLSX/`。

- 断点续采、批量文件处理、停止/恢复、实时进度、独立日志窗口
- 汇率自动获取（`open.er-api.com`）
- **跳图配置**：每个输入文件可单独设置需跳过的图片位置（如 `1,3`，`-1` 表示最后一张）
- **变体图进图库**：变体图片 URL 自动并入 `src_links`，确保 Shopify 导入时 `Variant Image` 与图库匹配
- **价格零值回落**：`compare_at_price` 为 `"0.00"` 时划线价回落为售价（`price2 = price1`），避免导入出现 0.00 划线价

> 早期的 `Shopify_JSON_GUI_SkipOption.py`（排除选项版）功能已合并进主线，现为 `SKIP_OPTIONS` 常量。

#### `shopify_scraper_workbench` — 加固版工作台

- 单文件版：`03CRAWLER/clean/shopify_scraper_workbench.py`（2587 行，归档）
- 模块化版：`03CRAWLER_Disassemble/shopify_scraper_workbench/`（9 模块）

由 `Shopify_JSON_GUI.py` 加固而来，业务逻辑零改动，差异集中在健壮性：

| 区域 | 加固内容 |
|---|---|
| `fetch_json` | 加 `stop_event` 支持，`tab.listen.wait` 改为 1 秒切片轮询，可即时响应停止 |
| 路径构造 | `build_output_file_path` / `build_output_folder_path` / `build_recovery_xlsx_path` / `_replace_path_component`，**保证输出路径绝不等于输入路径** |
| 采集主循环 | 拆出 `scraper_runner.py` 的 `ScraperRunner`（422 行）：文件遍历、按 `link-href` 续采、浏览器创建、重试、解析、`SAVE_EVERY` 自动保存与异常保护保存；通过 `stop_event` + `on_file_status` / `on_progress` / `on_pending_merge` 回调与 GUI 通信，**不导入 tkinter** |
| 保存 | `file_utils.py` 改为临时文件 + `os.replace` 原子写入（43–74 行） |
| 合并 | `_merge_worker` 统一走 `build_output_folder_path` |

`main.py` 因此从 1413 行减到 1175 行，只保留 GUI 与调度。**日常采集优先使用这个包**；`Shopify_JSON_GUI.py` 作为基线保留。

#### Scrapy 工作流

```bash
cd ShopifyJsonDetailPage
scrapy crawl {spider_name}
```

| 配置 | 值 |
|---|---|
| `SPIDER_LOADER_CLASS` | `utils.spider_loader.RecursiveSpiderLoader`（支持中文目录名递归加载） |
| `ROBOTSTXT_OBEY` | `False` |
| `CONCURRENT_REQUESTS_PER_DOMAIN` | `1` |
| `DOWNLOAD_DELAY` | `1` |
| `ITEM_PIPELINES` | `ShopifyJsonPipeline`（XLSX 输出 + 断点续采） |

Spider 按「域名+分类」中文子目录组织，每个源站一个 spider 文件（当前 5 个）。`utils/link_collector.py` 在 Spider 内调用 DrissionPage 完成链接采集。

| 特性 | GUI 工作流 | Scrapy 工作流 |
|---|---|---|
| 入口 | `03CRAWLER_Disassemble/shopify_scraper_workbench/main.py` | `scrapy crawl {spider}` |
| 链接采集 | `link_collector_pagination` / `detail_link_GUI.py` | `utils/link_collector.py`（Spider 内调用） |
| 架构 | 模块分离（`main.py` 只有 GUI） | 模块化，共享 `utils/` |
| 断点续采 | 支持 | Pipeline 支持 |
| 适用 | 日常批量，可视化进度 | 定制化爬虫开发 |

### 非 Shopify 站点采集

针对**不能直接使用标准 `{URL}.json`** 的站点：Shopify Headless、Next.js/RSC、Magento PWA、WooCommerce、React/Vue SPA、Shopline、自研电商等。

方法论详见 [`04非 Shopify 标准站点采集/非Shopify站点采集脚本开发流程.md`](DrissionPage_Json_Shopify/04非%20Shopify%20标准站点采集/非Shopify站点采集脚本开发流程.md)（1068 行）。

**核心原则**：只适配站点采集逻辑，最终输出主管线统一的 10 字段数据。

开发新站先解决三个问题——① 商品数据在哪里 ② 变体怎么对应价格和图片 ③ 页面什么时候真正加载完成。数据源选择优先级：

```
JSON / GraphQL / XHR  →  JSON-LD  →  Next.js/RSC 内嵌数据  →  渲染后商品 DOM  →  点击变体动态获取
```

硬性约束：不从商品名猜颜色或尺码、不人为创建站点不存在的变体、错误页/骨架页/空数据不算成功商品。

**公共工具** `non_shopify_utils.py`（187 行）：

| 函数 | 作用 |
|---|---|
| `fetch_exchange_rates()` | 汇率获取 |
| `convert_price_to_usd(price_str, currency, rates)` | 货币换算 |
| `format_price(value, rates, currency)` | 价格格式化 |
| `format_image_url(url, size="600x600", use_query_marker=False)` | 图片 URL 改写（支持追加 `?_600x600=1` 尺寸标记，不改写文件名） |
| `clean_body_html(body_html)` + `CleanBodyHtmlParser` | 详情 HTML 清洗 |

脚本置于 `非 Shopify 站点/` 子目录，通过 `sys.path.insert(0, parent)` 导入公共工具；**禁止** `from Shopify_JSON_GUI import ...` 或把 `03CRAWLER` 加入 sys.path——`04` 目录必须能独立运行。

**已接入站点**：

| 脚本 | 站点架构 | 数据源与特殊处理 |
|---|---|---|
| `snitch_crawler_GUI.py` (389) | Shopify + XPath | 渲染页 XPath 提取（`SIZE_XPATH`/`PRICE_XPATH`/`DESC_XPATH`/`LDJSON_XPATH`）；`IMAGE_EXCLUDE_KEYWORDS` 过滤 logo/促销图；**仅采集**，不做价格匹配与 WP 转换；每采 1 条即保存 |
| `dobell_crawler_GUI.py` (503) | Magento PWA | 颜色拆成独立页面，有尺码选项时 `styles1` 用 `Size#...`，不凭商品名猜颜色 |
| `joyfit_crawler_GUI.py` (502) | Shopline | `/products/*.json` 返回 403；数据全在 SSR HTML——ld+json `@type=Product` + `<script name="variant-data">`（价格为分）+ 媒体画廊 `li[data-media-id]` 做颜色→图映射 |
| `teeshoppen_crawler_GUI.py` (534) | Shopify Hydrogen | 标准 `.json`/`.js` 返回 Hydrogen HTML；改用 SSR 的 JSON-LD + DOM |
| `hespokestyle_crawler_GUI.py` (546) | WordPress + WooCommerce | 无商品 schema、无 `data-product_variations`；币种固定 USD 无需换算；仅 `select#jacket-size` 一个尺码维度，`onhand="0"` 缺货尺码须排除；图片走 JSON-LD `ImageObject`，NitroPack 懒加载图是占位符故用 `use_query_marker=True` |

### styles 单脚本线（Data logic）

`DrissionPage_Json_Shopify/Data logic/` 是一条**无 GUI、无命令行参数**的独立加工线：把已有的 10 字段 styles 表一步跑完「原价修正 → 价格匹配 → 未匹配统计 → Shopify 转换 → WP 转换」五步（详见[价格匹配 B-1](#b-styles-表匹配data-logic-线较新)的规则）。

```bash
# 唯一配置项：脚本内的 INPUT_FILE 常量
python "DrissionPage_Json_Shopify/Data logic/Styles_递增价格匹配.py"
```

- 输出写在**输入文件同级**的 `原价/` 与 `价格匹配/` 两个子目录
- 不写 xlsx 备份/日志文件，过程信息全部打印到 stdout
- `styles/lovedandco_styles.xlsx` 及其一次运行产物留作对照样例

定位是**基准/实验工具，不是生产入口**：`INPUT_FILE` 硬编码、无测试、仓库内除 Preposition README 的一句引用外没有任何代码依赖它。它的价值在于定义了「递增匹配」这套规则，Preposition 包即按此对齐而来。

### 转换与辅助工具

#### `Utils/switch_gui.py` — 8 合 1 转换工具（1505 行）

左侧工具导航（`_create_nav_button`，每行 4 个）+ 右侧可拖拽 PanedWindow 运行日志。除图片转换外均支持文件夹递归。

| # | 工具 | 输入 | 说明 |
|---|---|---|---|
| 1 | WP 价格检查 | wp-*.csv | 对比 `PRICE_LIBRARY`，列出缺失价位 |
| 2 | WP 还原 styles | wp-*.csv | **反向转换**，WP 格式回退到 styles xlsx |
| 3 | SKU 查重 | wp-*.csv | 检测重复 SKU 与孤立变体（Parent 无对应父商品） |
| 4 | SKU 更新 | wp-*.csv | 批量重写 SKU（父 SKU 由 Handle 的 MD5 决定，保证可复现） |
| 5 | Shopify 价格修正 | Shopify csv / 目录 | `Sale price == Regular price` 时把 Regular 上调 ×1.2 |
| 6 | Excel 按 title 拆分 | styles xlsx / 目录 | 按分组列归组、组内按价格升序后轮询分成 N 份 |
| 7 | 图片批量转 JPG | 图片目录（+可选 csv） | WebP/BMP/TIFF/GIF/AVIF/HEIC → JPG，可同步改写 CSV 的 `Images` 后缀 |
| 8 | 图片链接精简 | styles xlsx / 目录 | 按 `styles1` 精简 `src_links`：被引用的图片全留，未引用的只留前 N 张（默认 4） |

内置 `PRICE_LIBRARY`（**215 个**预设价位，8.99–599）。功能 5 / 6 / 8 覆盖原文件时会先写临时文件再替换，或先备份为 `.bak`，避免中途失败损坏数据。

#### `Utils/Single function/` — 单功能脚本（6 个）

b439ea5 把散在 `Utils/` 根下的脚本统一中文命名并归入子目录（`equal_price_fix.py`→`价格修正.py`、`to_jpg.py`→`图片转JPG.py`、`split_by_title.py`→`按标题拆分Excel.py`、`Drawing lots.py`→`按SKU抽样.py`、`extract_img_local.py`→`提取本地图片.py`，另有新增的`描述字段清洗.py`）。多数是 `switch_gui` 对应功能的脚本版，路径写在文件顶部的常量里。

| 脚本 | 行数 | 用途 |
|---|---|---|
| `价格修正.py` | 392 | 遍历 CSV，`Sale price == Regular price` 时把 `Regular price` 改为 1.2 倍（`MARKUP = 1.2`），输出后缀默认 `_价格修正`；双击开 GUI，也支持 `python 价格修正.py <文件或目录> [--suffix 后缀] [--no-overwrite]` |
| `图片转JPG.py` | 290 | WebP/BMP/TIFF/GIF/AVIF/HEIC → JPG，自动 `pip install Pillow`，可选 `pillow-heif`；联动改写关联 CSV 的图片后缀 |
| `按标题拆分Excel.py` | 369 | 按 `title`（可改列）分组、组内按 `price1` 升序，轮询拆成 N 份（默认 2），输出 `源名_1.xlsx`、`源名_2.xlsx`；GUI |
| `按SKU抽样.py` | 70 | 从 `wp-*.csv` 的 `Type=variable` 主体商品里随机抽 `SAMPLE_N` 个（默认 10），连同其变体导出 `源名_抽测数据.csv`（UTF-8 无 BOM）；配置项为脚本顶部的 `SRC` 路径 |
| `描述字段清洗.py` | 52 | 删除 `Description` 里的 `B/C/BC + 两位数字`编号与 `Celebrity Favorites such as` 整段 `<p>`，压缩空白；**原地写回**，先备份 `.bak` |
| `提取本地图片.py` | 73 | 解析 `styles1` 中 `@` 后的图片链接，与 `src_links` 比对后精简（`switch_gui` 功能 8 的最简版） |

---

## 模块化拆分版

`03CRAWLER_Disassemble/` 存放巨型单文件的**界面与业务分离**重构版，四个目录：`link_collector_pagination/`（链接采集）、`shopify_scraper_workbench/`（商品采集加固版）、`standard_collection_process/`（标准采集流水线，唯一带测试的版本）、`standard_collection_process_Preposition/`（前者的同源副本，多出域名替换、关键词剔除与百分位分档价格匹配）。

拆分本身是**零行为变更的纯搬家**，采用 AST 机械提取保证函数体逐字节一致。此后的修复以独立提交进入模块化版本，`03CRAWLER/`（含 `clean/`）的单文件版**不自动同步**——这正是当前两套实现行为分叉的原因（见[已知问题与修复状态](#已知问题与修复状态)）。

### `link_collector_pagination/`（6 模块）

| 模块 | 行数 | 职责 |
|---|---|---|
| `config.py` | 133 | UI 样式 + 采集参数 + 浏览器参数 + `CategoryItem` / `CollectConfig` 数据类 |
| `logger.py` | 139 | 多 sink 日志广播 + `log_exc` + stdout 重定向/还原 + `TextRedirector` 兜底 |
| `browser_utils.py` | 171 | `create_browser` / `interrupt_page` / `run_js_with_retry` / `get_outer_html` / `evaluate_xpath` |
| `file_utils.py` | 394 | `load_existing_domain_rows`（三态返回）/ `build_recovery_file_path` / `update_link_index` / `SaveReport` / `save_results`（原子写入） |
| `collector.py` | 430 | `parse_categories` / `build_page_url` / `extract_product_links` / `extract_links_by_xpath` / `collect_page_links` / `CollectorRunner`（`on_save_report` 回调） |
| `main.py` | 493 | **仅 GUI**（唯一导入 tkinter 的模块）；`_on_close` 等待保存 + `CLOSE_WAIT_TIMEOUT=60` 兜底 |

该目录另有 `README.md`（258 行），是按当前实现撰写的操作手册，含快速使用、输出文件、限制与排查。

### `shopify_scraper_workbench/`（9 模块）

| 模块 | 行数 | 职责 |
|---|---|---|
| `config.py` | 149 | 全部常量（含 `PRICE_LIBRARY` / `SHOPIFY_COLUMNS` / `WP_HEADERS` / `MAX_STYLES_LENGTH`）+ 全局可变配置 `KEEP_POSITIONS` / `SKIP_POSITIONS` |
| `logger.py` | 150 | `TextRedirector` + sink 管理 + stdout 还原（tkinter 延迟导入） |
| `browser_utils.py` | 43 | `create_browser` + 代理常量 |
| `network_utils.py` | 119 | `to_shopify_json_url` / `group_targets_by_json_url` / `fetch_exchange_rates` / `convert_price_to_usd` / `fetch_json` |
| `product_parser.py` | 464 | 28 个解析函数 + `CleanBodyHtmlParser` + `with_image_filter` 上下文管理器 + `InvalidProductError`（含 `MAX_STYLES_LENGTH` 校验） |
| `file_utils.py` | 144 | Excel 读写（**已改原子写入**）+ 4 个安全路径构造 + recovery |
| `scraper_runner.py` | 422 | `ScraperRunner`：采集主循环，回调通信，**禁止导入 tkinter**（9c4e94e 新增） |
| `converter.py` | 487 | 17 个转换函数（Shopify 转换 / 价格匹配 / WP 导出） |
| `main.py` | 1175 | **仅 GUI**（`ScraperApp`），实例化 `ScraperRunner` 并接回调 |

> 该目录的 `README.md`（350 行）**最后更新于 ed8c447（2026-09-17），早于 `scraper_runner.py` 拆分**：其模块表未列 `scraper_runner`、仍称「调度在 `ScraperApp._run_task()`」、仍称「Excel 保存尚无原子替换」，与当前代码矛盾，阅读时以源码为准。

### `standard_collection_process/`（54 文件）

采集流水线的分层实现：

| 层 | 文件 | 要点 |
|---|---|---|
| `crawler/` | `browser.py`(26) `links.py`(107) `product_api.py`(32) | 浏览器与请求 |
| `processing/` | `product.py`(115) `options.py`(95) `style_codec.py`(64) `variants.py`(66) `images.py`(126) `pricing.py`(52) `shopify.py`(173) `woocommerce.py`(164) `currency.py`(79) `html_cleaner.py`(66) `cells.py`(9) | 解析、Styles 编解码、A 线价格匹配 |
| `storage/` | `workbooks.py`(110) `merge.py`(38) `exports.py`(67) `paths.py`(60) `tables.py`(49) `checkpoints.py`(17) | 读写、合并、导出 |
| `pipeline/` | `controller.py`(139) `link_task.py`(156) `product_task.py`(194) `conversion_task.py`(48) `models.py`(82) | 状态机与任务 |
| `ui/` | `app.py`(648) `components.py`(527) `log_window.py`(114) `theme.py`(72) | Tk 界面 |
| `resources/` | `columns.py`(97) `price_library.py`(219) | 列定义与价格库（**WP 26 列**） |
| `tests/` | 8 个测试文件 55 项 + `fixtures/reference.json` | 见[运行方式与测试](#运行方式与测试) |

`config.py` 关键值：`LINK_MAX_RETRIES=5`、`RETRY_DELAY=5`、`PAGE_LOAD_TIMEOUT=30`、`DELAY=0.5`、`PRODUCT_MAX_RETRIES=5`、`RETRY_BASE_DELAY=30`、`SAVE_EVERY=20`、`SKIP_OPTIONS=None`、`MAX_STYLES_LENGTH=32000`、`PROXY_URL="http://127.0.0.1:7897"`。被判为无效的商品（JSON 结构校验失败或 `styles1` 超长）逐个跳过，不影响同一文件中的其他商品。

三阶段操作与人工断点（该包 README 的 mermaid 流程图）：链接采集 → 商品 JSON 采集 → 选择商品数据文件夹做「合并并转换」，输出 `xlsx/`+`csv/` 三态与 WP CSV。

### `standard_collection_process_Preposition/`

同源副本，`processing/product.py`(115) 与 `pipeline/product_task.py`(194) 与主版本逐字节一致，**改一处需同步另一处**。差异集中在：

| 文件 | 行数 | 差异 |
|---|---|---|
| `processing/domain.py` | 47 | 合并时按规则替换来源域名 |
| `processing/exclude.py` | 30 | 按关键词剔除商品（整体字面量匹配、大小写不敏感） |
| `processing/pricing.py` | 808 | **百分位分档价格匹配**（见 B-2） |
| `processing/shopify.py` | 194 | 主版本 173 行的对齐扩展 |
| `processing/woocommerce.py` | 149 | 主版本 164 行的对应实现 |
| `pipeline/conversion_task.py` | 61 | 主版本 48 行 |
| `storage/merge.py` / `exports.py` | 59 / 72 | 主版本 38 / 67，输出改为 `原价/`、`价格匹配/` 两子目录 |
| `config.py` | 40 | 多一个 `EXCLUDE_NAME_KEYWORDS = ["Gift Card"]` |

该包 README（439 行）明确记载：价格匹配已换成百分位分档，与 `Data logic/` 脚本不再是同一算法；并承认匹配结果与行序相关。

---

## 运行方式与测试

### GUI 工具

```bash
# 链接库
python "Assembly line/jsonl_gui.py"

# 链接采集（无限滚动）
python DrissionPage_Json_Shopify/03CRAWLER/detail_link_GUI.py

# 模块化拆分版：目录名以数字开头，不是合法包名，采用平铺绝对导入 +
# sys.path.insert(0, dirname(abspath(__file__)))，必须 cd 进目录后运行
cd DrissionPage_Json_Shopify/03CRAWLER_Disassemble/shopify_scraper_workbench
python main.py

# 这两个包用包内 requirements（pandas 3.0.3），在 03CRAWLER_Disassemble 下运行
cd DrissionPage_Json_Shopify/03CRAWLER_Disassemble
pip install -r standard_collection_process/requirements.txt
python standard_collection_process/main.py
```

拆分版根目录名以数字开头（`03CRAWLER_Disassemble`），因此 `standard_collection_process` 的模块以**顶层包名**方式导入，`sys.path` 需指向 `03CRAWLER_Disassemble/`。

### 测试（55 项，已实测）

`.venv` 中**没有 pytest**，测试是 `unittest` 风格且通过 `from common import ...` 依赖 tests 目录本身在 `sys.path` 上，正确跑法：

```bash
cd DrissionPage_Json_Shopify/03CRAWLER_Disassemble/standard_collection_process/tests
PYTHONPATH="<repo>/DrissionPage_Json_Shopify/03CRAWLER_Disassemble" \
  <repo>/.venv/Scripts/python.exe -m unittest discover -s . -p "test_*.py"
```

2026-09-29 实测：`Ran 55 tests in 3.767s — OK`。

| 测试文件 | 项数 | 覆盖 |
|---|---|---|
| `test_pipeline.py` | 13 | 停止/保存流程、续采、无效商品跳过、`MAX_STYLES_LENGTH` 撞限 |
| `test_processing.py` | 13 | 规格转义、尺码识别、价格规则、图片过滤 |
| `test_ui_and_architecture.py` | 6 | Tk 界面与分层约束（业务模块不导入 tkinter） |
| `test_storage.py` | 6 | 读写与保存失败恢复 |
| `test_alignment.py` | 5 | 与单文件基线的一致性（对照 `fixtures/reference.json`） |
| `test_conversion.py` | 5 | Shopify / WP 转换 |
| `test_baseline_rules.py` | 4 | 基线数据规则 |
| `test_folder_merge.py` | 3 | 文件夹合并 |

`link_collector_pagination/`、`shopify_scraper_workbench/`、`standard_collection_process_Preposition/` **不含测试目录**，拆分期历史测试未随源码保留，旧 README 里的 131 / 350 项目前**不可复现**——改动这三个目录没有回归网。Preposition 与主版本共享的两个文件只有一次性验证。

拆分后业务模块可独立单测，无需启动 GUI。

> 端到端对比需固定 `random.seed`（Handle/SKU 由 random 生成，A 线价格匹配还按 Handle 分组顺序分配），比对时剔除 Handle/SKU/Barcode/MPN 等随机列。

---

## 输入输出目录组织

```
01INPUT_XLSX/
└── A0916/                                  # 批次日期（当前还有 A0730_2 A0801 A0812 A0901 A0910）
    ├── 周晓东/                             # 操作员
    │   ├── halcys.com周晓东内衣家居服/      # 目标站点 + 分类
    │   │   ├── {source_domain}.xlsx        #   产品链接文件
    │   │   └── links_index.json            #   分类 URL 索引
    │   ├── 换clarityxs.com女装/             # 「换」= 换源站点重采
    │   ├── 换astrilas.com女装/
    │   └── 换ulkri.com/
    ├── 林文健/
    │   ├── 换solsticeva.com大码女装/        # 「补」/「减量」等前缀同为任务标记
    │   └── covecradle.com服装/
    └── 陈港/
        ├── 补delfias.com礼服/
        └── sylvelin.com女装/
```

`02OUTPUT_XLSX/` 结构相同，每个分类目录下包含：

| 产物 | 说明 |
|---|---|
| 各来源站原始 xlsx | 阶段一采集结果（10 字段） |
| `{分类}_合并.xlsx` | 「合并并转换」产物 |
| `xlsx/` + `csv/` 子目录 | A 线三态：原价 / 价格匹配 / 未匹配到的价格 |
| `wp-{分类}_合并/` | WooCommerce 导入 CSV |
| `..._价格匹配_价格修正.csv` | 经 `Utils/Single function/价格修正.py` 处理后的版本 |
| `原价/` `价格匹配/` | B 线（styles 表匹配）产物结构 |

已完成批次的数据归档至 `00Historical data/`。

---

## 文档索引

| 文档 | 内容 | 注意 |
|---|---|---|
| [`03CRAWLER/styles数据处理流程报告.md`](DrissionPage_Json_Shopify/03CRAWLER/styles数据处理流程报告.md) | 1107 行。四阶段流水线逐行详解：采集 → Shopify 转换 → 价格匹配 → WP 转换，含 SKU 生成规则、完整数据流示例、关键全局变量、最终产物结构 | 按 **A 线（Shopify 表匹配）** 口径撰写，不含 B 线 |
| [`Data logic/Data logic.md`](DrissionPage_Json_Shopify/Data%20logic/Data%20logic.md) | 226 行。`Styles_递增价格匹配.py` 的五步详解、输入输出目录、与 GUI 版差异对照、关键常量、已知坑 | 两处已过期：称 variable 父 SKU 为 `SKU{time}{line}`（脚本实际用 Handle 的 MD5）；称 WP 输出 27 列含 `Tags`（当前实现为 26 列不含） |
| [`04非 Shopify 标准站点采集/非Shopify站点采集脚本开发流程.md`](DrissionPage_Json_Shopify/04非%20Shopify%20标准站点采集/非Shopify站点采集脚本开发流程.md) | 1068 行。非标准站点接入方法论：核心原则、公共工具、12 步开发流程、页面特征识别表、10 字段契约、完整性校验 | — |
| [`03CRAWLER_Disassemble/link_collector_pagination/README.md`](DrissionPage_Json_Shopify/03CRAWLER_Disassemble/link_collector_pagination/README.md) | 258 行。拆分版结构、功能范围、快速使用、输出与限制 | 与当前代码基本一致 |
| [`03CRAWLER_Disassemble/shopify_scraper_workbench/README.md`](DrissionPage_Json_Shopify/03CRAWLER_Disassemble/shopify_scraper_workbench/README.md) | 350 行。拆分版结构、循环导入打破方案、改进窗口、限制与排查 | **陈旧**（2026-09-17，早于 `scraper_runner.py` 与原子写入），三处与代码矛盾，见上文说明 |
| [`03CRAWLER_Disassemble/standard_collection_process/README.md`](DrissionPage_Json_Shopify/03CRAWLER_Disassemble/standard_collection_process/README.md) | 359 行。三阶段操作、数据规则、价格与导入格式、停止与异常、代码结构、55 项测试范围 | — |
| [`03CRAWLER_Disassemble/standard_collection_process_Preposition/README.md`](DrissionPage_Json_Shopify/03CRAWLER_Disassemble/standard_collection_process_Preposition/README.md) | 439 行。同上，另含商品排除关键词、域名替换、百分位分档价格匹配、与 Styles 脚本的一致性验证 | 已注明与 Data logic 脚本分叉 |

`03CRAWLER_Disassemble/*.md` 四份逐文件源码归档已于 2026-09-21（b439ea5）删除，不再需要「改源码同步归档」。

---

## 已知问题与修复状态

拆分过程中梳理出的高优先级缺陷。**「已修」只指模块化版本**，`03CRAWLER/`（含归档 `clean/`）不会自动同步。

| 问题 | 位置 | 状态 |
|---|---|---|
| `{h: (v or "")}` 把数值 `0`、`False` 静默写成空串 | `standard_collection_process{,_Preposition}/storage/workbooks.py`、`clean/standard_collection_process.py` | **已修**（改为 `v if v is not None`） |
| 同上 | `shopify_scraper_workbench/file_utils.py`、`03CRAWLER/Shopify_JSON_GUI.py` | **仍在**（前者自身 README 限制章节仍列此条） |
| 同上 | `link_collector_pagination/file_utils.py:76`（`str(row[0] or "")`） | 该文件已无 `load_xlsx`，改走 `load_existing_domain_rows`，同类写法残留 1 处 |
| XLSX 写入非原子，中断会损坏原文件 | `link_collector_pagination/file_utils.py`、`shopify_scraper_workbench/file_utils.py` | **已修**（临时文件 + `os.replace`，9c4e94e） |
| 同上 | `03CRAWLER/Shopify_JSON_GUI.py`、`clean/shopify_scraper_workbench.py` | **仍在**（直接 `wb.save`；`clean/standard_collection_process.py` 已原子） |
| 全量重写会丢失旧文件的额外列 | `shopify_scraper_workbench/file_utils.py`（仅重建固定 `CSV_FIELDS`） | **仍在**；`clean/standard_collection_process.py` 额外保留隐藏页 `_crawl_config` |
| 后缀替换用 `split(".")[-1]`，路径含点号时行为异常 | `shopify_scraper_workbench/converter.py` 及三个单文件版各 1 处 | **仍在** |
| Handle/SKU 由 `random` 生成，非确定性 | `converter.py`、两个拆分包的 `processing/shopify.py`、`Shopify_JSON_GUI.py`、Data logic 脚本 | **仍在**（`switch_gui` 功能 4 的 MD5 父 SKU 是唯一的确定性方案） |
| `styles1` 超长静默截断 | 见 [Styles 编码格式](#styles-编码格式) | 模块化三处已加 `MAX_STYLES_LENGTH=32000`；单文件版**未加** |
| B-2 分档匹配结果与行序相关 | `standard_collection_process_Preposition/processing/pricing.py` | 已在该包 README 记录，未修 |
| `SKIP_OPTIONS` 在单文件版为 `["ships from"]`、两个拆分包为 `None` | 各 `config.py` | 实现分叉，按站点需要确认 |

---

## 本次修订要点

相对 2026-09-17 版 README，依据当前代码更正/补充（核对日期 2026-09-29）：

1. **目录树重画**：新增 `03CRAWLER/clean/`（三个单文件版归档，a703afb 纯移动、无引用）、`DrissionPage_Json_Shopify/Data logic/`、`Utils/Single function/`（6 个中文命名脚本）；`03CRAWLER/` 根目录现仅剩 `Shopify_JSON_GUI.py`、`detail_link_GUI.py` 和流程报告。
2. **价格匹配拆成独立章节**：原文「75%–125% 容差 / 最近值 / 1.2 倍 / 每价位一次」只适用于 A 线；新增 B 线（Data logic `[0.6,1.1]` 递增、Preposition 五档百分位 + usage 计分 + 仅 CSV 输出）与差异速查表。
3. **删除失效的「源码归档」章节**：四份 `03CRAWLER_Disassemble/*.md` 已删除（b439ea5）。
4. **模块表更新**：`shopify_scraper_workbench` 由 8 模块改 **9 模块**（新增 `scraper_runner.py` 422 行，`main.py` 1413→1175）；`link_collector_pagination` 各模块行数刷新并补充原子写入、`SaveReport`、关窗等待三项已落地能力。
5. **修正计数与常量**：`detail_link_GUI.py` 1156→**1193**（原数字在写下时即为旧值）；`links.jsonl` 1881→**1979** 条并补批次/状态实况；`WP_HEADERS` 区分 27（含 `Tags`）与 26（拆分包）；流程报告 1078→**1107** 行、非 Shopify 方法论 **1068** 行。
6. **测试章节改为可执行命令**：`.venv` 无 pytest，给出 `unittest discover` + `PYTHONPATH` 的正确跑法，2026-09-29 实测 **55 项通过**，并列出 8 个测试文件的项数分布。
7. **Preposition 包从「同源副本」升级为有独立差异表**：域名替换、`EXCLUDE_NAME_KEYWORDS=["Gift Card"]`、`pricing.py` 808 行分档实现、`原价/`+`价格匹配/` 输出结构。
8. **「已知待修项」改为带状态的表**：逐条标注已修/仍在，并指明「已修」只在模块化版本，单文件版不自动同步。
