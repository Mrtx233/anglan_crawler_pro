# Shopify 商品采集工作台

`shopify_scraper_workbench` 读取 Excel 中的商品链接，通过 DrissionPage 请求 Shopify 商品 JSON，解析商品名称、价格、变体、图片地址和描述，按输入文件分别保存商品中间表。

本项目的流程到中间表保存结束。合并、价格匹配及 Shopify / WooCommerce 导出由独立的 `Xlsx_processing_platform` 完成。

## 1. 三个项目的衔接

```text
link_collector_pagination
    分类页 → title/link 链接表
             ↓
shopify_scraper_workbench（本项目）
    链接表 → 请求商品 JSON → 商品中间表
             ↓
Xlsx_processing_platform
    读取中间表文件夹 → 合并 → Shopify 转换 → 价格匹配 → WooCommerce 导出
```

三个程序分别启动，通过磁盘上的 Excel 文件衔接，不会自动启动彼此。本工作台不包含“合并并转换”按钮，也不上传店铺或下载图片文件。

## 2. 环境与启动

需要带 Tkinter 的 Python、Chromium 系浏览器，以及 `DrissionPage`、`openpyxl`、`certifi`。本项目不再依赖 pandas；独立 Excel 处理平台仍需要 pandas。

在当前 macOS 项目目录运行：

```sh
cd /Users/malyu/Downloads/PythonCode/Anglan_Crawler_Pro_Max/shopify_scraper_workbench
python3 -m pip install DrissionPage openpyxl certifi
python3 main.py
```

可用 `python3 -m tkinter` 检查 Tkinter。Windows 中如果解释器命令是 `python`，将上述命令中的 `python3` 替换为 `python`。

当前目录没有依赖锁定文件，未固定最低 Python、浏览器和依赖版本组合。

### 浏览器代理

代理配置位于 `browser_utils.py`：

```python
PROXY_SERVER = "http://127.0.0.1:7897"
```

不需要该代理时设为 `None` 或 `""`，程序便不传入浏览器代理参数。修改源码配置后重新启动程序。

汇率请求使用 Python `urllib`，不直接复用 Chromium 的代理设置；其连接取决于 Python 的代理环境。

## 3. 输入与输出目录

### 输入表格式

输入为 `.xlsx`，使用活动工作表，第一行跳过，第 2 行起按列位置读取：

| 第 1 列：分类 title | 第 2 列：商品 link |
|---|---|
| Women > Dresses | https://example.com/products/dress |
| Women > Tops | https://example.com/products/top?variant=123 |

- 按前两列的位置读取，不按表头名称定位；第 3 列及以后不参与采集。
- 分类和链接转为字符串并清除首尾空白；空链接跳过，分类允许为空。
- 输入单元格中的 `0`、`False` 会按空值处理。
- 只扫描文件夹第一层的 `.xlsx`，按文件名排序；跳过以 `~` 开头的文件。
- 不扫描子目录，不接受 CSV、`.xls` 或直接粘贴的商品 URL。
- `links_index.json` 不参与处理；分类来自输入 Excel，而不是 JSON 索引。
- 不要把恢复文件、商品中间表或转换结果混入输入目录；工作台不会自动识别并排除这些 `.xlsx`。

### 输出路径

路径中的第一个完整 `01INPUT_XLSX` 段会被替换为 `02OUTPUT_XLSX`，匹配时忽略大小写。没有该目录段时，在输入文件所在文件夹下创建 `02OUTPUT_XLSX` 子目录。

```text
Anglan_Crawler_Pro_Max/
├── 01INPUT_XLSX/
│   └── demo/
│       ├── example.com.xlsx
│       └── another.example.xlsx
└── 02OUTPUT_XLSX/
    └── demo/
        ├── example.com.xlsx
        └── another.example.xlsx
```

例如：

```text
/data/01INPUT_XLSX/demo/a.xlsx → /data/02OUTPUT_XLSX/demo/a.xlsx
/data/demo/a.xlsx              → /data/demo/02OUTPUT_XLSX/a.xlsx
```

每个输入文件对应一个同名商品中间表。代码会检查输出路径，阻止直接覆盖输入文件。仓库根目录的 `01INPUT_XLSX/`、`02OUTPUT_XLSX/` 已通过 `.gitignore` 排除，不纳入正常 Git 提交。

## 4. 界面操作

1. 点击“浏览文件夹”，选中链接表所在文件夹，程序自动读取；也可手工填写路径后点击“读取文件夹”。手工更改路径后需重新读取。
2. 核对文件列表及分类。每个文件下显示按首次出现顺序去重的分类名。
3. 根据需要填写文件行或分类行的“保留图片”“跳过图片”。无筛选需求时全部留空。
4. 点击“开始采集”。任务运行期间，路径和图片规则输入被禁用。
5. 点击“查看日志”或“查看运行日志”，检查采集、重试和保存信息。日志窗口可独立隐藏，主任务继续执行。
6. 如需停止，点击“停止任务”，等待保存完成；正常结束后也应核对日志中的失败数和输出文件。

分类行可单独折叠，也可用“全部折叠/全部展开”统一操作。折叠仅隐藏控件，不清除规则。重新读取文件夹会重建列表和规则输入。

界面保留操作按钮、输入示例和状态，不再显示额外操作说明段落。图片规则未持久化到配置文件，重新启动后需重新填写。

## 5. 图片规则

### 位置含义

| 设置 | 结果 |
|---|---|
| 文件行两栏均空 | 使用全部图片 |
| 保留 `1,2` | 仅保留实际 position 为 1、2 的图片 |
| 跳过 `1,3` | 排除实际 position 为 1、3 的图片 |
| 保留 `-1` | 保留最后一个 position 对应的图片 |
| 跳过 `-1,-2` | 排除最后两个 position 对应的图片 |

负数基于排序、去重后的 position 列表倒数定位。支持英文逗号、中文逗号及全角减号 `－`；不支持 `1-3` 范围写法。`0` 不是“全部”的标记。

### 文件默认与分类覆盖

| 分类行状态 | 实际使用 |
|---|---|
| 两栏都留空 | 整套继承文件行规则 |
| 任意一栏非空 | 完全使用分类行的两栏，空的那栏不过滤 |
| 某栏填 `-` | 该栏显式不过滤，并启用分类独立规则 |

例如文件保留 `1,2`，分类只填写跳过 `1`，该分类会使用“不过滤保留范围、跳过 1”的规则，不继承文件的保留限制。

**`-` 作为显式不过滤标记只在分类行生效。** 文件行直接使用整数列表解析，文件保留栏填 `-` 会形成空保留列表并过滤全部图片；文件行需要不过滤时请留空。

两栏同时设置时，保留优先，跳过不生效。非整数片段通常会被忽略，界面不做严格校验：保留 `abc` 会得到空保留列表，删除全部图片；仅跳过 `abc` 则等于不跳过任何图片。越界保留位置也可能得到空图片集合。

规则同时作用于商品图片列表和变体关联图，实际生效规则会记录在日志中。规则只作用于本轮新采商品，已有输出会被续采跳过；更改规则不会重新处理历史记录。

## 6. 请求、去重与续采

### URL 处理

请求地址去掉查询参数、fragment 和末尾斜杠，再追加 `.json`；已有 `.json` 后缀不重复追加：

```text
https://example.com/products/dress?variant=123#info
→ https://example.com/products/dress.json
```

| 场景 | 行为 |
|---|---|
| 同一输入文件中多个 URL 对应同一 JSON 地址 | 只请求一次，保留首条的分类和来源 URL |
| 不同输入文件中出现相同商品 | 没有全局去重，分别处理 |
| 对应旧输出存在非空 link-href | 将该来源 URL 视为已完成，跳过请求 |

写出的 `link-href` 去掉 query 和 fragment，但不会统一去掉末尾斜杠或 `.json`。续采比较这个字符串，不使用 JSON 地址作匹配键，所以不同链接形式可能导致重复采集。

### 浏览器与重试

每个有待采商品的文件创建一个隐身浏览器并复用标签页；处理完该文件后关闭。无商品或全部已完成的文件不启动浏览器。

请求先监听网络响应状态，再读取页面 `innerText` / `textContent`，必要时读取 `pre` 文本解析 JSON。监听等待约 30 秒，并每秒检查停止信号；这不是整个请求的硬性超时，页面加载本身仍可能阻塞。

每商品初次请求之外最多重试 5 次，最多共 6 次请求：

| 情况 | 额外重试前的等待 |
|---|---|
| HTTP 429 | 30、60、90、120、150 秒 |
| 其他状态或未取得可用 JSON | 5、10、15、20、25 秒 |

重试轮次及商品之间另有 0.5 秒等待，等待可被停止事件打断。超过重试次数后记为失败，继续下一商品。

`styles1` 超长引起的 `InvalidProductError` 会跳过该商品并计失败；其他未处理的浏览器、解析或文件异常可能中止整个任务，后续文件不会自动继续。

当前没有严格校验返回 JSON 是否包含完整商品对象；HTTP 200 或日志显示成功不代表字段一定完整。

## 7. 保存、停止和异常处理

- 读取对应历史输出后，保留旧行并追加新商品；不自动刷新已有价格、描述或图片。
- 每个文件的新成功数达到 20 的正整数倍时触发自动保存。当前条件按成功计数判断，若停在该倍数后连续失败，可能重复保存同一批结果。
- 文件结束、停止或异常清理时，有结果便尝试保存当前文件；浏览器关闭失败不会直接跳过保存。
- 使用同目录临时文件写入完整 Excel，再通过 `os.replace` 替换目标。原子替换失败时保留原文件并清理临时文件。
- 旧 Excel 读取失败时，禁止覆盖旧表，本轮另存 `文件名.recovery-YYYYMMDD-HHMMSS.xlsx`；同名时追加序号。恢复文件不会自动合回原表。
- 写入失败不会自动生成 recovery；收尾保存失败会记日志，但界面仍可能显示“已完成”。必须结合保存日志和实际输出确认成功。

点击“停止任务”后设置停止事件；当前未提交商品会被放弃，此前结果尝试保存。停止操作不会直接中断浏览器的 `tab.get()`。

采集中关闭主窗口会先确认，再请求停止并以 200 毫秒间隔等待工作线程，最多等待 60 秒。超时后强制关闭，可能丢失未落盘结果。关闭独立日志窗口只是隐藏日志，不会停止任务。

进度条展示当前文件待采商品的处理进度。最终成功/失败数只统计本轮处理；总数据行数包含历史行。部分商品失败时文件仍可能标记“已完成”，不代表全部商品成功。

## 8. 中间表字段与数据变化

每个商品一行，固定写入以下 10 列，内容经过解析与清洗，不是原始 JSON 备份：

| 字段 | 内容 |
|---|---|
| title | 输入链接表中的分类文本 |
| name | product.title，基本原样取值 |
| price1 | 参考变体售价，经换汇及格式化逻辑处理 |
| price2 | 非零 compare_at_price；缺失或为零时回退到售价 |
| styles1 | 选项、变体组合、价格、关联图片的内部编码 |
| styles2 | 空字符串 |
| styles3 | 空字符串 |
| src_links | 用 # 连接的图片地址 |
| link-href | 去掉 query/fragment 的来源 URL |
| details | 清洗后的 body_html |

保存会重建工作簿，手工增加的列、样式及其他工作表不保留。读取历史输出使用 `v or ""`，数值 `0` 和 `False` 会变为空字符串。

### 8.1 价格与汇率的实际执行顺序

参考变体优先选择 `position=1`，不存在时使用第一个变体。

每次任务开始调用汇率获取函数，来源为 `https://open.er-api.com/v6/latest/USD`。同一进程已有非空缓存便复用，不逐商品更新。

`format_price()` 当前按以下顺序执行：

1. 读取变体的 `price_currency`。有汇率和币种时，尝试将原始数值除以该币种的 USD 基准汇率，返回两位小数字符串；USD、缺少对应汇率或无法解析金额时保留原值。
2. 对上一步结果进行单位判断：含小数点则直接按金额解析，不含小数点则按整数除以 100。
3. 格式化输出，最多保留两位小数并去除末尾多余的零。

例如未执行换汇时，`"1999" → "19.99"`，`"1999.00" → "1999"`，`"19.90" → "19.9"`。

**币种和汇率不能决定原始金额的单位。当前代码仍使用小数点判断单位，尚未改成按接口约定统一单位后换汇。** 若原始整数已经表示金额，会被误除以 100；若整数表示分且先成功换汇，又可能因换汇结果带小数点而漏除以 100。例如假设 `1 USD = 7 CNY`，原始 `1999` 若表示分，应为约 `2.86 USD`，当前顺序会得到 `285.57 USD`。

仅当变体带 `price_currency` 且有可用汇率时才能换汇。程序不会另行查询店铺币种，中间表也没有独立币种列；不能将全部输出价格视为已统一美元。缺币种或汇率时仍会执行后续格式化，而不是完全原样保存。

### 8.2 选项与变体

- 按 option 的 position 排序，排除名称包含 `ships from` 的选项，匹配时忽略大小写。
- 保留有值的选项，包括单值选项。
- 名称含 size，或全部值都属于配置中的尺码集合时，选项名统一为 `Size`。
- 名为 Type 的选项改为 `Style`；Color 优先排到第一位。
- 根据选项值生成笛卡尔积，先找完整匹配变体，再找第一选项匹配变体，最后回退到参考变体。
- 可能生成源站实际不存在的组合，且不按库存可用性过滤售罄变体。

`styles1` 示例：

```text
Color#Red&Size&S$19.99$24.99@https://cdn.example/red_600x600.jpg#Blue&Size&M$21.99$25.99
```

`#` 分隔主选项名称与组合，`&` 分隔选项字段，`$` 分隔价格，`@` 引出图片。选项名和值中的反斜杠、`&`、`#`、`@` 使用反斜杠转义；`$` 不在转义集合中。

编码超过 `MAX_STYLES_LENGTH=32000` 字符时，整个商品跳过，记录失败，不写入中间表。此限制只针对 styles1，不是所有单元格的统一长度检查。

### 8.3 图片地址

图片按 position 排序并根据规则过滤，在文件扩展名前插入 `_600x600`，保留查询参数和 fragment。例如 `a.jpg?v=1` 变为 `a_600x600.jpg?v=1`。

变体关联图优先按 Color 建立映射，没有 Color 则按第一选项；同一选项值取首次找到的可用关联图。关联图未出现在主图片列表时追加，已有相同地址不重复追加。这不是对所有原始图片的完整去重。

输出只保存远程 URL，不下载、不验证可访问性，也不检查站点是否支持该尺寸后缀。

### 8.4 商品描述

当前清洗规则为：

- 删除整个 `<a>…</a>` 元素，包括链接文字、内部标签和图片。
- 删除所有 `<img>` 标签，包括自闭合形式；链接之外的正文保留。
- 移除其他标签的属性，包括 style、class 等。
- 反复移除空成对标签。
- 对文本进行 HTML 转义，移除码点大于 `0xFFFF` 的字符，例如部分 emoji。

例如：

```html
<!-- 原始描述 -->
<p class="intro">介绍<a href="https://example.com"><strong>点击购买</strong></a><img src="a.jpg">正文</p>

<!-- 清洗后 -->
<p>介绍正文</p>
```

处理使用 HTMLParser，不是浏览器级的 HTML 修复或完整安全过滤。未闭合的 a 标签可能使其后文本也被跳过，应核对结构异常的原始描述。

## 9. 后续合并与导出

1. 确认本工作台已保存各文件商品中间表。
2. 启动同级项目 `Xlsx_processing_platform/main.py`。
3. 填写中间表所在文件夹，例如 `02OUTPUT_XLSX/demo`，点击“合并并转换”。
4. 在该文件夹的 `processing_results/` 查看合并表、Shopify 原价与匹配价，以及 WooCommerce CSV。

处理平台每次重新读取磁盘文件，因此手工修改后的中间表可以参与处理。不要将第一阶段仅有 title/link 的链接表直接交给处理平台。

## 10. 模块与配置

| 模块 | 职责 |
|---|---|
| main.py | GUI、扫描、图片规则输入、线程启动、状态与关窗等待 |
| scraper_runner.py | 文件遍历、续采、请求重试、解析与保存 |
| browser_utils.py | Chromium 创建、代理与随机端口 |
| network_utils.py | URL 规范化、文件内去重、汇率及 JSON 请求 |
| product_parser.py | 商品字段、价格、变体、图片和描述解析 |
| file_utils.py | Excel 读写、输出路径、原子保存及 recovery 命名 |
| config.py | 采集参数、字段、图片规则、汇率缓存及 UI 配置 |
| logger.py | 日志输出及 stdout/stderr 转发 |

| 配置位置 | 项目 | 默认值 |
|---|---|---|
| config.py | DELAY | 0.5 秒 |
| config.py | MAX_RETRIES | 5 次额外重试 |
| config.py | RETRY_BASE_DELAY | 429 等待基数 30 秒 |
| config.py | SAVE_EVERY | 20 次新成功 |
| config.py | MAX_STYLES_LENGTH | 32000 字符 |
| config.py | SKIP_OPTIONS | ships from |
| config.py | IMAGE_FILTER_NONE | 分类行显式不过滤标记 `-` |
| config.py | INPUT_DIRNAME / OUTPUT_DIRNAME | 01INPUT_XLSX / 02OUTPUT_XLSX |
| browser_utils.py | PROXY_SERVER | http://127.0.0.1:7897 |
| browser_utils.py | PORT_RANGE / START_RETRIES | 9222～9322 / 3 次 |
| main.py | CLOSE_WAIT_TIMEOUT / CLOSE_POLL_INTERVAL_MS | 60 秒 / 200 毫秒 |

部分配置在模块导入时绑定，修改后重启。图片规则通过全局配置临时切换，不适合多个解析线程并发使用。

### Python 调用

在本项目目录运行以下示例，将路径替换为实际输入文件：

```python
import threading
from scraper_runner import ScraperRunner

stop_event = threading.Event()
task_files = [{
    "path": "/data/01INPUT_XLSX/demo/example.com.xlsx",
    "keep": "",
    "skip": "",
    "categories": {
        "Women > Dresses": {"keep": "1,2", "skip": ""},
    },
}]
runner = ScraperRunner(
    task_files=task_files,
    task_folder="/data/01INPUT_XLSX/demo",
    stop_event=stop_event,
    on_file_status=lambda index, text, state: print(index, text, state),
    on_progress=lambda current, total, text: print(current, total, text),
)
rows = runner.run()  # 阻塞执行，内部尝试保存；未处理异常会向外抛出
print(len(rows))
# 其他线程可调用 stop_event.set() 请求停止
```

返回行数包含历史行，不等于本轮新增数，也不能单独证明落盘成功。直接调用 `parse_product()` 时需自行处理 `InvalidProductError`。

三个项目有同名顶层模块（如 config、file_utils），应分别在独立进程和各自目录运行，避免模块缓存冲突。

## 11. 排查与验证范围

| 现象 | 检查方向 |
|---|---|
| 浏览器无法启动或页面打不开 | DrissionPage、浏览器、代理地址、调试端口 |
| 200 响应但字段为空或报错 | 实际 JSON 是否为完整商品结构 |
| 长时间停在同一商品 | 日志中的 429 重试、网络监听和页面加载 |
| 改图片规则后结果没变 | 旧输出是否已让该商品被续采跳过 |
| 图片全部消失 | 保留规则是否无效、越界，文件保留栏是否误填 `-` |
| 商品数量少 | 请求失败、styles1 超长、文件内去重或任务异常 |
| 金额异常 | 原始单位、price_currency、汇率，以及先换汇后格式化的顺序 |
| 描述中的链接文字消失 | 当前规则会删除整个 a 元素，是预期行为 |
| 出现 recovery 文件 | 旧输出无法读取，需人工核对并处理恢复结果 |
| 已完成但文件没更新 | 查看保存失败日志、文件占用及目录权限 |

当前目录没有依赖锁定文件或随附自动测试套件。本轮开发曾进行模块导入、模拟采集/保存/续采及描述清洗检查；这些离线检查不代表在线站点兼容性或店铺导入已验证。

实际验收建议用独立输入输出目录，覆盖少量商品、重复 URL、文件/分类图片规则、第二次续采、停止保存、损坏旧表 recovery 和带 a/img 的描述。日志默认只在 GUI 或标准输出展示，不自动保存日志文件。
