# Shopify 商品采集工作台

`shopify_scraper_workbench` 是基于 Tkinter 的桌面批处理工具：读取 Excel 中的商品链接，用 DrissionPage 访问 Shopify 商品 JSON，解析商品、价格、变体、图片和描述，保存中间表，再由用户触发合并、Shopify 格式转换、价格匹配与 WooCommerce CSV 导出。

它与 `link_collector_pagination` 组成两阶段流程：

```text
分类列表页 → 链接采集器 → title/link Excel
          → 商品工作台 → 商品中间表
          → 手动“合并并转换” → Shopify 原价 / 匹配价 → WooCommerce CSV
```

本工具不直接上传商品到店铺，不下载图片文件，也不自动发现分类。本文依据当前源码描述行为，不沿用拆分历史中的测试通过承诺。

## 1. 主要功能

- 读取所选文件夹第一层的 `.xlsx`，按文件名排序依次处理。
- 将商品 URL 规范化为 `.json` 请求地址，在每个输入文件内去重。
- 根据已有输出的 `link-href` 跳过已采集商品，实现文件级续采。
- 每个输入文件可独立配置保留或跳过的图片位置。
- 提取商品名称、参考变体价格、变体组合、图片 URL 和清洗后的 HTML 描述。
- 对 HTTP 429 与其他失败执行不同等待策略，支持停止与自动保存。
- 手动合并本次任务数据，导出 Shopify XLSX/CSV、价格匹配表、未使用价格清单和 WooCommerce CSV。

适用前提是商品详情 URL 对应的 `.json` 地址可访问，并返回含 `product` 对象的预期结构。当前不包含登录、验证码处理或任意网页 DOM 商品解析。

## 2. 环境与启动

需要带 Tkinter 的 Python、可由 DrissionPage 启动的 Chromium 系浏览器，以及 `DrissionPage`、`openpyxl`、`pandas`。当前目录没有依赖锁定文件，最低版本与完整兼容范围未固定。

在项目目录中执行：

```powershell
python -m pip install DrissionPage openpyxl pandas
python main.py
```

Tkinter 可通过 `python -m tkinter` 检查。当前机器上的项目目录为：

```text
D:\A_PythonCode\anglan_crawler_pro\DrissionPage_Json_Shopify\03CRAWLER_Disassemble\shopify_scraper_workbench
```

浏览器代理在 **`browser_utils.py`** 中配置，默认 `PROXY_SERVER = "http://127.0.0.1:7897"`。设为 `None` 或 `""` 时不向浏览器传入代理参数。修改后重新启动程序。

汇率请求使用 Python `urllib`，不会直接复用这里的 Chromium 代理配置；它取决于 Python 所使用的代理环境。

## 3. 输入文件要求

### Excel 格式

每个输入文件必须是 `.xlsx`，活动工作表第 1 行为表头，第 2 行开始为数据：

| title | link |
|---|---|
| Women > Dresses | https://example.com/products/sample-dress |
| Women > Tops | https://example.com/products/sample-top?variant=123 |

- 第 1 列是分类文本，第 2 列是商品链接；按列位置读取，不按表头名称查找。
- URL 为空的行跳过，分类可为空。
- 第 3 列及以后的输入内容不会用于采集。
- 仅扫描所选文件夹第一层，不递归；以 `~` 开头的临时文件不处理。
- GUI 不接收 CSV、`.xls` 或单独粘贴的 URL。
- 链接采集器生成的域名 Excel 可直接使用；`links_index.json` 不参与工作台处理。

不要把采集输出、合并结果、recovery 文件与正常输入混在同一待扫描目录，否则它们也可能被当成输入表。

### 推荐目录布局

```text
D:\data\
├── 01INPUT_XLSX\
│   └── demo\
│       ├── example.com.xlsx
│       └── another.example.xlsx
└── 02OUTPUT_XLSX\
    └── demo\                  # 工具自动建立
```

输出路径优先将路径中的第一个完整 `01INPUT_XLSX` 段替换为 `02OUTPUT_XLSX`，忽略大小写。如果不存在该名称，则在输入文件所在目录下创建 `02OUTPUT_XLSX` 子目录；不是在原文件上覆盖。

例如：

```text
D:\data\01INPUT_XLSX\demo\a.xlsx → D:\data\02OUTPUT_XLSX\demo\a.xlsx
D:\data\demo\a.xlsx              → D:\data\demo\02OUTPUT_XLSX\a.xlsx
```

## 4. GUI 操作流程

1. 选择分类文件夹；通过浏览按钮选择时会自动扫描，也可填写路径后点击“读取文件夹”。
2. 核对文件列表，为每个文件填写图片保留/跳过规则，留空表示不过滤。
3. 点击“开始采集”。任务运行时文件夹和图片规则输入被禁用。
4. 打开日志窗口查看请求、失败重试、自动保存和输出路径。关闭日志窗口不会关闭主工作台。
5. 如需中止，点击“停止”，等待当前流程退出并保存，确认日志后再关闭主窗口。
6. 采集结束后检查各文件输出，再点击“合并并转换”。该步骤不会自动执行。
7. 在输出目录的 `xlsx/`、`csv/` 和 `csv/wp-.../` 中检查转换产物。

进度条主要展示当前文件的待采商品处理进度。结束汇总中的成功/失败统计对应本次请求，数据总行数还包含历史续采记录；“已完成”状态也不保证所有商品都成功，应结合失败数量和日志判断。

**检查表格与编辑表格不同：**“合并并转换”使用内存中的采集结果，不会重新读取你刚手工修改的输出 Excel。若需要转换修改后的表格，请使用后文的独立转换入口。

## 5. 图片筛选规则

| 输入 | 含义 |
|---|---|
| 保留、跳过均留空 | 使用全部图片 |
| 保留 `1,2` | 仅保留 `position` 为 1、2 的图片 |
| 跳过 `1,3` | 排除 `position` 为 1、3 的图片 |
| 保留 `-1` | 仅保留最后一个位置的图片 |
| 跳过 `-1,-2` | 排除最后两个位置的图片 |

- 正数匹配 JSON 图片的实际 `position` 字段；负数根据排序去重后的 position 列表从后向前计算。
- 支持英文/中文逗号和全角减号 `－`。不支持 `1-3` 这样的范围表达式。
- **保留规则优先于跳过规则**，两者同时填写时按保留规则执行。
- 越界保留位置可能得到空图片集合；不要使用 `0` 表示“全部”。
- 非整数片段会被忽略，GUI 不会严格报错。例如保留框填入 `abc` 可能形成空保留列表，过滤掉全部图片，应填写明确的整数列表。
- 过滤同时作用于主图片列表和变体关联图片。输出的是加上 `_600x600` 尺寸后缀的远程图片 URL，原查询参数会保留；不会下载或验证图片是否可用。
- 已有商品会被续采跳过，修改图片规则不会重新处理这些记录。需要重新采集时，先备份并移走对应输出文件，或改用新的输出目录结构。

## 6. 请求、去重、重试与续采

### URL 规范化

`network_utils.to_shopify_json_url()` 去掉 URL 的查询参数和 fragment，移除末尾斜杠，再追加 `.json`；已经以 `.json` 结尾的不再追加。

```text
https://example.com/products/dress?variant=123#info
→ https://example.com/products/dress.json
```

同一输入文件中规范化到同一 JSON URL 的记录只请求一次，保留首次出现的分类与来源 URL。不同输入文件之间没有全局去重，合并时也不会再次按商品 URL 去重。

### 浏览器与请求

每个有待采商品的输入文件创建一个隐身浏览器，文件处理结束后关闭；已全部完成或无商品的文件不会启动浏览器。

请求通过浏览器网络监听取得状态码，再读取页面 `innerText` / `textContent`，必要时读取 `pre` 文本解析 JSON。监听等待最多约 30 秒、按 1 秒切片检查停止信号；这不是整个请求的硬性 30 秒上限，`tab.get()` 本身也可能阻塞。

### 失败重试

- 每商品初次请求之外最多重试 `MAX_RETRIES=5` 次，即最多 6 次请求。
- 429 的重试等待依次为 30、60、90、120、150 秒。
- 其他状态或无法取得有效 JSON 时，等待依次为 5、10、15、20、25 秒。
- 重试轮次及商品间还有 `DELAY=0.5` 秒等待；等待可被停止事件中断。
- 达到次数上限后记录失败，继续下一个商品。
- 单个商品被解析拒绝（`parse_product()` 抛 `InvalidProductError`，当前唯一来源是 styles1 超过 `MAX_STYLES_LENGTH`）时只跳过该商品并记录原因，`item` 保持为空，按普通失败计数；同一文件的其他商品继续采集。
- **其他未处理的异常不同于普通失败：**浏览器异常等可能结束当前文件并终止整个任务；当前实现不保证自动继续后续文件。

### 续采与保存

1. 读取对应旧输出表，收集非空 `link-href` 作为完成集合。
2. 跳过来源 URL 已在完成集合中的商品，只请求剩余商品。
3. 保留历史行，追加新结果；新成功数达到 20 的倍数时自动保存，并在文件结束、停止或异常清理时尝试保存。
4. 旧输出无法读取时，原文件不覆盖，本轮另存 `文件名.recovery-YYYYMMDD-HHMMSS.xlsx`，重名追加序号。

续采按来源 URL 字符串比较，不校验内容是否完整，也不自动刷新价格、标题和图片。输入中使用 `.json` 与普通详情页等不同 URL 形式，可能影响历史完成匹配。recovery 不会自动并回正常输出文件。

正常结束或停止且有数据时，生成内存中的 `pending_merge`。异常退出路径可能无法建立该待合并状态。程序重启后它不会恢复；可重跑同一输入以读取旧结果，或者调用独立转换入口。

## 7. 商品中间表字段与解析规则

每个商品一行，固定写入 10 列：

| 字段 | 内容 |
|---|---|
| `title` | 输入表中的分类文本 |
| `name` | `product.title` 商品名称 |
| `price1` | 参考变体售价 |
| `price2` | 非零 `compare_at_price`；缺失或为零时回退到售价 |
| `styles1` | 主要选项、变体组合、价格和关联图片的内部编码 |
| `styles2`、`styles3` | 当前采集器写空，转换器保留读取能力 |
| `src_links` | 使用 `#` 连接的图片 URL |
| `link-href` | 去掉 query/fragment 的来源 URL |
| `details` | 清洗后的 `body_html` |

### 价格与汇率

- 参考变体优先取 `position=1`，不存在则取第一个变体。
- `format_price()` 将没有小数点的输入按分除以 100，有小数点的输入按金额处理，例如 `1999 → 19.99`、`"19.99" → 19.99`。源数据整数若本身已是元，会被按分处理，需要核对。
- 每次任务开始调用汇率获取函数；进程中已有非空缓存则直接复用。来源为 `https://open.er-api.com/v6/latest/USD`。
- 仅在变体带 `price_currency` 且有对应汇率时尝试换算 USD；缺币种、汇率缺失或请求失败时保留未换算输入，不能把所有输出价格视为已成功转换美元。
- 当前价格格式化与换汇有先后顺序上的约束，尤其应核对“整数分值 + 非 USD 币种”的实际金额；中间表没有独立币种列。

### 选项与变体

- 按 option 的 `position` 排序，排除名称包含 `ships from` 的选项。
- 仅保留多值选项；单值选项不作为变体维度。
- 尺码选项统一名为 `Size`；名为 `Type` 的选项改为 `Style`；`Color` 优先排到第一位。
- 对选项值生成笛卡尔积，优先匹配完整变体，其次匹配第一选项，最后使用参考变体。
- 因此可能生成源站未实际提供的组合，且不按库存可用性筛掉售罄变体。转换后须核对组合数量。

`styles1` 的典型格式为：

```text
Color#Red&Size&S$19.99$24.99@https://cdn.example/red_600x600.jpg#Blue&Size&M$21.99$25.99
```

`#` 分隔主选项名与各组合，`&` 分隔选项字段，`$` 分隔价格，`@` 引出图片。选项名/值中的反斜杠、`&`、`#`、`@` 使用反斜杠转义；`$` 不在当前转义集合中。该格式是工具内部约定，不是 Shopify 原生字段。

`styles1` 长度上限为 `config.MAX_STYLES_LENGTH`（32000 字符）。超过上限时 `parse_product()` 抛 `InvalidProductError`，商品整体跳过，不写入来源表和 `pending_merge`，日志记录实际长度与上限。取 32000 而非 Excel 单元格上限 32767，是为边界问题留出余量。此规则与 `standard_collection_process` 同步。

**注意调用方：**`parse_product()` 现在会抛 `InvalidProductError`。直接调用它的代码需要自行捕获，否则异常会向上冒泡；`main.py` 的采集循环已就地捕获并跳过该商品。

### 描述清洗

去掉 `<a>` 标签但保留文本，去掉 `<img>`，移除标签属性和空成对标签，对文本进行 HTML 转义，并移除码点大于 `0xFFFF` 的字符（例如部分 emoji）。这属于内容整理，不是完整的 HTML 安全过滤器。

## 8. 合并与导出产物

假设输入文件夹名为 `demo`，则正常转换的目录结构为：

```text
02OUTPUT_XLSX/demo/
├── example.com.xlsx
├── another.example.xlsx
├── demo_合并.xlsx
├── xlsx/
│   ├── demo_合并_Shopify_原价.xlsx
│   ├── demo_合并_Shopify_价格匹配.xlsx
│   └── demo_合并_Shopify_未匹配到的价格.xlsx
└── csv/
    ├── demo_合并_Shopify_原价.csv
    ├── demo_合并_Shopify_价格匹配.csv
    ├── demo_合并_Shopify_未匹配到的价格.csv
    └── wp-demo_合并/
        └── wp-demo_合并_Shopify_原价.csv
```

未使用价格清单只在仍有价格未被使用时生成。CSV 使用 UTF-8 BOM（`utf-8-sig`）。同名转换产物会重写。

### Shopify 转换

`converter.styles_to_shopify()` 读取中间表，展开变体和图片，生成 `config.SHOPIFY_COLUMNS` 约定的 49 列原价表：

- 每商品生成随机后缀 Handle，每变体生成随机 SKU；同一次转换内去重，重跑不会保持相同标识。
- `name → Title`，`details → Body (HTML)`，`title → Collection/Vendor`，`link-href → Link-Href`。
- 每商品输出行数取变体组合数、图片数与 1 的最大值；部分行可能仅用于附加图片。
- 默认已发布、库存 999、库存跟踪 `shopify`、库存策略 `deny`、履约 `manual`、需要运输、不计税、重量 0、单位 kg。
- 这些默认值是代码填充规则，不是源站真实库存、税务或物流数据。文件包含项目自定义字段，实际导入前需要按使用的导入流程核对。

### 价格匹配

原价表保留采集阶段的结果；价格匹配表会改价：

1. 按 Handle 分组，各组内有效变体按原售价升序处理。
2. 在 `PRICE_LIBRARY` 中寻找原价 75%～125% 范围内、全文件尚未使用的价格。
3. 选择满足条件的**最小价格**，不是最接近原价的价格。
4. 更新 `Variant Price`，将 `Variant Compare At Price` 设为匹配价的 1.2 倍，并新增“匹配价格”列。
5. 没有候选价的行保留原价；每个库价格在整个文件中最多使用一次。

“未匹配到的价格”表列出的是**价格库尚未使用的价格**，不是匹配失败的商品清单。检查失败商品可查看价格匹配表中“匹配价格”为空且原售价有效的行。

### WooCommerce 导出及当前自动链路限制

WP 导出按 Handle 聚合，生成 `simple` 或 `variable` 父商品及 `variation` 子行，转换分类、图片、属性和价格，默认库存 99999。售价和原价均大于零且售价更高时，将原价提升到售价。

`WP_HEADERS` 配置含 27 项，但写出时排除了 `Tags`，所以实际 CSV 为 26 列。父 SKU 根据 Handle 派生；由于上游 Handle 随机，整条流水线重跑后的标识仍可能不同。

**当前自动链路通常只生成原价 WP 文件。** 价格匹配 XLSX 位于 `xlsx/`，而配套 CSV 在 `csv/`；代码却只替换 XLSX 路径的扩展名去找 CSV，仍在 `xlsx/` 查找，找不到时静默跳过价格匹配 WP 导出。需要匹配价 WP 时，可显式调用：

```python
from converter import shopify_to_wp

shopify_to_wp(
    r"D:\data\02OUTPUT_XLSX\demo\csv\demo_合并_Shopify_价格匹配.csv",
    print,
)
```

该调用会在对应 `csv/wp-demo_合并/` 中生成 `wp-demo_合并_Shopify_价格匹配.csv`。自动 WP 转换异常会记录日志，但不会阻止外层返回；不能仅凭最后的“转换完成”判断所有目标文件均已生成。

## 9. 模块结构和配置

| 模块 | 职责 |
|---|---|
| `main.py` | `ScraperApp` GUI、文件扫描、采集调度、续采、停止、合并线程 |
| `config.py` | 采集参数、图片配置、汇率缓存、字段表、价格库与 UI 样式 |
| `browser_utils.py` | 代理、随机端口和 Chromium 创建 |
| `network_utils.py` | URL 规范化、分组去重、汇率和浏览器 JSON 请求 |
| `product_parser.py` | 价格、选项、变体、图片、描述和中间表解析；styles1 长度上限与 `InvalidProductError` |
| `file_utils.py` | Excel 读写、安全输出路径、recovery 命名 |
| `converter.py` | Shopify 转换、价格匹配与 WooCommerce 导出 |
| `logger.py` | GUI 日志 sink、输出重定向与恢复 |

业务解析与转换函数可以独立调用；完整采集调度仍在 `ScraperApp._run_task()` 中，并未抽成独立任务运行器。汇率缓存和图片过滤保存在 `config` 中，`with_image_filter()` 在退出时恢复配置，但不是线程隔离配置，不适合多个解析线程同时切换。

| 配置位置 | 项目 | 默认值 / 说明 |
|---|---|---|
| `config.py` | `DELAY` | 0.5 秒 |
| `config.py` | `MAX_RETRIES` | 5 次额外重试 |
| `config.py` | `RETRY_BASE_DELAY` | 429 线性等待基数 30 秒 |
| `config.py` | `SAVE_EVERY` | 每 20 次新成功触发保存条件 |
| `config.py` | `KEEP_POSITIONS` / `SKIP_POSITIONS` | 默认 `None`；GUI 按文件临时覆盖 |
| `config.py` | `SKIP_OPTIONS` | `["ships from"]`，名称包含匹配 |
| `config.py` | `MAX_STYLES_LENGTH` | 32000，styles1 超过即跳过该商品 |
| `config.py` | `INPUT_DIRNAME` / `OUTPUT_DIRNAME` | `01INPUT_XLSX` / `02OUTPUT_XLSX` |
| `config.py` | `PRICE_LIBRARY` | 价格匹配候选池 |
| `config.py` | `CSV_FIELDS` / `SHOPIFY_COLUMNS` / `WP_HEADERS` | 中间表及导出字段约定 |
| `browser_utils.py` | `PROXY_SERVER` | `http://127.0.0.1:7897` |
| `browser_utils.py` | `PORT_RANGE` / `START_RETRIES` | `(9222, 9322)` / 3 次启动尝试 |

部分配置在模块导入时绑定，修改源码后应重启。`build_output_folder_path()` 的非约定路径分支仍写死 `02OUTPUT_XLSX`，修改 `OUTPUT_DIRNAME` 时还需核对该分支。

## 10. 独立调用转换器

在项目目录内执行或保存为同目录脚本，可绕过 GUI 转换已有中间表：

```python
from converter import styles_to_shopify, price_match, shopify_to_wp

# 完整转换：Shopify 原价、价格匹配，并尝试自动 WP 导出。
original_xlsx = styles_to_shopify(
    r"D:\data\02OUTPUT_XLSX\demo\demo_合并.xlsx", print
)
print(original_xlsx)

# 对已有 Shopify 表可单独调用：
# matched_xlsx = price_match(r"D:\data\shopify.xlsx", print)
# wp_csv = shopify_to_wp(r"D:\data\shopify.csv", print)
```

`styles_to_shopify()` 实际返回原价 Shopify XLSX 路径，尽管其函数注释称返回匹配价路径；`price_match()` 返回价格匹配文件路径，`shopify_to_wp()` 返回 WP CSV 路径。

转换必需列为 `title`、`name`、`price1`、`price2`、`details`、`src_links`、`styles1`、`styles2`、`styles3`；建议保留 `link-href` 方便追溯。该入口会写文件并执行完整转换，适合使用已确认的中间表。

两个配套项目使用同名顶层模块（如 `config`、`file_utils`），应分别在各自目录和进程运行，避免 Python 模块缓存冲突。

## 11. 当前限制与排查

| 现象或限制 | 说明与处理 |
|---|---|
| 浏览器无法启动或页面打不开 | 检查 DrissionPage、浏览器、代理和端口；代理配置在 `browser_utils.py` |
| 请求一直失败或返回 200 仍失败 | 核对 `.json` 地址和返回正文，200 页面也可能不是 JSON；当前没有严格商品结构校验 |
| 长时间停在一个商品 | 429 会逐步延长等待，查看日志；停止等待可响应，但浏览器加载不一定立即中断 |
| 修改图片或价格规则后结果不变 | 旧输出按 `link-href` 跳过；备份旧表后重新采集 |
| 行数少于预期，日志出现「styles1 长度…超过上限」 | 该商品组合数过多，编码超长被整体跳过；属预期行为，不是请求失败 |
| 合并中重复商品 | 去重仅在单个输入文件内，跨文件和合并阶段不去重 |
| 修改了输出表但合并结果没变化 | GUI 合并使用内存数据；对编辑后的中间表调用独立转换器 |
| 重新打开工具不能合并 | `pending_merge` 只在内存中；重跑读取历史数据，或独立调用转换器 |
| 某文件异常后后续文件没执行 | 当前未处理异常会终止任务；查看首个异常并修正后重跑（单个商品被 `InvalidProductError` 拒绝不属于此类，不会中断文件） |
| 数值 0 或 False 变为空 | `load_xlsx()` 使用 `v or ""`，会将这些值读为空串 |
| 手工附加列、格式或工作表丢失 | 保存仅重建固定 10 列，额外信息不保留；人工备注宜存放在独立文件 |
| recovery 反复出现 | 原输出仍无法读取；recovery 不自动替代或合并回原输出 |
| 匹配价 WP 文件缺失 | 自动路径查找目录不一致，见第 8 节显式转换示例 |
| 关闭窗口后未完整保存 | 关窗发送停止信号后立即销毁 GUI，未等待 daemon 工作线程；应先停止并等保存完成 |

Excel 保存使用直接写入，尚无原子替换；recovery 仅保护“读不出旧表”的情况，不能防止写到一半时进程退出。合并转换线程也没有独立停止/关窗等待机制，执行转换时应等待结束。日志默认仅在 GUI/标准输出中展示，不自动保存日志文件。

当前目录未附带 `tests/`、依赖锁定文件或测试入口。旧 README 的等价性、GUI 和主流程测试数量不代表当前可复现的验证结果。验收宜在独立数据目录中覆盖：少量商品采集、重复链接、保留/跳过图片、第二次续采、停止保存、损坏旧表 recovery，以及原价/匹配价/WP 产物检查。本文不声明已完成在线采集或店铺导入验证。
