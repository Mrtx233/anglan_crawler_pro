# Excel 数据处理平台

填写商品中间表所在文件夹，点击「合并并转换」，依次完成：

```text
文件夹中的商品 Excel → 字段规范化 → 合并 → Shopify 原价 → 价格匹配 → WooCommerce 原价及匹配价 CSV
```

## 启动

需要带 Tkinter 的 Python，以及 openpyxl、pandas；无需浏览器、代理或 DrissionPage。

```sh
python3 -m pip install openpyxl pandas
python3 main.py
```

## 输入

选择 `shopify_scraper_workbench` 保存商品中间表的文件夹，例如 `02OUTPUT_XLSX/补delfias.com礼服`。无需再次采集、读取文件列表或填写输出路径。每次点击都重新读取磁盘文件，手工修改过的表格会参与本次处理。

只读取文件夹第一层的 `.xlsx`，按文件名排序。不递归读取子目录；跳过临时/隐藏文件、`_合并.xlsx`、名称含 `_Shopify_` 的导出表和 `.recovery-` 恢复文件。

活动工作表必须有以下列（按列名匹配，列顺序不限）：

```text
title, name, price1, price2, styles1, styles2, styles3, src_links, details
```

`link-href` 可选，缺失时补空。第一步链接采集器的 `title/link` 表不能直接作为输入。缺少必需列、重复列名或损坏文件会中止任务并显示文件信息，不静默漏掉数据。全空行跳过，数值 0 和 False 在合并时保留。

## 输出

自动写入所选文件夹的 `processing_results/`，源文件保持原样：

```text
demo/
├── a.xlsx
├── b.xlsx
└── processing_results/
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
            ├── wp-demo_合并_Shopify_原价.csv
            └── wp-demo_合并_Shopify_价格匹配.csv
```

未使用价格清单仅在价格库仍有未使用值时生成。再次运行会重写同名结果，自动生成的子目录不会被扫描回输入。

## 数据处理规则

- 合并按文件顺序追加，保留固定 10 列，不按商品链接去重，不保留额外列、样式和其他工作表。
- 此平台不重新请求商品，不做汇率换算、描述清洗或图片位置筛选，使用中间表已有内容。
- Shopify 转换解码 styles 字段并展开变体、图片，生成 49 列；Handle 和变体 SKU 每次随机生成。
- 保留原有默认值：已发布、库存 999、库存跟踪 shopify、需要运输、不计税、重量 0、单位 kg。
- 价格匹配按原价的 75%～125% 找到价格库中尚未使用的最小值；每个库价格全文件最多使用一次。匹配后划线价为匹配价的 1.2 倍，没有候选价则保留原价。
- 「未匹配到的价格」指价格库中未使用的价格，不是失败商品清单。
- WooCommerce 导出保留父商品/变体结构，默认库存 99999，输出 26 列（排除 Tags）。售价和原价均为正且售价高于原价时，将原价提高到售价。
- 原价和匹配价都导出 WooCommerce CSV，任一转换失败会报告失败，不显示全部完成。

合并表采用临时文件加原子替换；转换产物沿用直接写入。运行中关闭窗口会等待任务完成，失败则留在窗口供检查日志。转换失败时可能已生成部分文件，可修正后重新运行。

## 模块

| 文件 | 职责 |
|---|---|
| main.py | 单路径 GUI、后台任务及日志 |
| processing_runner.py | 扫描、校验、合并及转换调度 |
| file_utils.py | 中间表读取与原子保存 |
| converter.py | Shopify、价格匹配及 WooCommerce 导出 |
| style_utils.py | styles 字段解析 |
| config.py | 字段、价格库及输出目录配置 |

可脱离 GUI 调用：

```python
from processing_runner import process_folder
report = process_folder('/path/to/02OUTPUT_XLSX/demo')
print(report.output_dir)
```

与其他两个项目分别在各自目录和进程运行，避免同名模块冲突。

## title 字段规范化

合并前调用独立模块 `field_rules/title.py`。原始中间表不修改，规范后的 title 写入合并表；name 也会按独立规则处理，尚未设置规则的其他字段保持读取值。

- 必填字符串，默认最长 100 字符、最多两级（单级也允许），允许英文字母和数字。
- 每个单词按 Title Case 规范；缩写白名单默认包含 USB。可通过 TitleRules.whitelist 添加标准拼写，例如 iPhone。
- 首尾普通空格删除，连续普通空格压缩，层级分隔统一为 ` > `。
- 合并前自动清理禁止标点和 emoji：替换为空格后压缩，避免词语粘连；换行、制表符和特殊空白转为普通空格，不可见控制字符删除。中文等非英文文字仍需人工确认，不直接删除文字。
- 保留半角 `>` 的层级含义；全角 `＞` 作为禁止符号清理，不自动推断层级。空层级仍会报错。
- 数字分数为斜杠例外：`3 / 4` 统一为 `3/4`，也允许独立分数标题。分母不能为零，普通文本中的斜杠自动清理。
- 超长或超出层级数时报错，不截断、不丢弃商品。层级上限、长度、数字开关和白名单集中在独立规则文件中。
- 日志记录原值、规范值、来源文件和真实 Excel 行号。任意 title 或 name 不合格时，本次不写合并表，也不执行后续转换；历史结果保持原样。

例如 ` women clothing>3 / 4 sleeve dresses ` 会写为 `Women Clothing > 3/4 Sleeve Dresses`；`Home & Garden` 变为 `Home Garden`；`Prom Dresses，` 变为 `Prom Dresses`；`Summer-Dress` 变为 `Summer Dress`。清理后为空的标题仍报错。

## name 字段规范化与品牌替换

独立规则位于 `field_rules/name.py`，默认将旧品牌替换为新品牌。

- 所选文件夹的最后一级名称必须包含一个明确域名。例如 `补delfias.com礼服` 提取目标品牌 `delfias`；不从父目录取值。
- 每条商品的旧品牌来自其来源文件名。例如 `babaroni.co.uk.xlsx` 提取 `babaroni`。来源文件需使用域名命名，不能继续使用上方结构示意中的 `a.xlsx`、`b.xlsx` 等占位名称；文件夹也不能仅叫 demo。
- 当前支持普通两段域名和代码中列出的常见复合后缀（包含 co.uk），支持 www 前缀。无法唯一识别、未知多段后缀或子域名时停止处理，提示人工确认，不猜测品牌。
- 匹配完整品牌词或来源完整域名，不区分大小写；优先替换域名，避免留下 co uk。名称不含旧品牌时不添加新品牌，只执行文本规范。
- 替换后执行每词首字母大写、其余小写，默认缩写白名单含 USB。新品牌 `delfias` 输出为 `Delfias`。
- 特殊符号、emoji 清理为空格，空格压缩；隐藏控制字符删除。name 不保留 title 的层级分隔符，`>` 同样清理。
- 数字分数 `3 / 4` 统一为 `3/4`；分母为零时报错。Unicode 分数如 ¾ 暂未确定转换规则，仍提示人工确认。
- name 必填且必须是字符串；清理后为空或含未支持的非英文文字时报错。商品名称长度暂未设业务上限，不沿用 title 的 100 字符限制。
- 规范后的值写入合并表，来源文件保持不变。title 或 name 校验失败时本批不写合并表、不启动转换，日志记录来源文件、行号和字段。

```text
文件夹：补delfias.com礼服
来源文件：babaroni.co.uk.xlsx
原名称：BABARONI Elegant 3 / 4-Sleeve Dress!
合并名称：Delfias Elegant 3/4 Sleeve Dress
```

## price1 / price2 字段规范化

分别由 `field_rules/price1.py`、`field_rules/price2.py` 管理，在合并前处理：

- 接受数字及数字字符串，使用 Decimal 十进制计算，按 ROUND_HALF_UP 四舍五入到两位小数。
- 合并表写入数字类型，并设置 Excel 单元格格式 `0.00`，例如 12 显示为 12.00。
- 先规范 price1，再规范 price2；按规范后的两位小数比较。price2 小于 price1 时，改为 price1 × 1.1，再四舍五入到两位小数；相等或更大时保留规范值。
- 允许零；空值、负数、布尔值、NaN、无穷大及不能解析的价格报错，阻止本批输出，不自动补零或推测币种。
- 日志记录字段、来源行号及前后数值，源中间表不修改。

示例：price1=`19.995`、price2=`18`，合并表保存为数字 20.00 和 22.00。

本轮仅规范中间表的 price1/price2 两列。styles1 内部编码的变体价格尚未制定规则，不在此步骤改写；后续 Shopify 转换仍可能优先使用 styles1 的变体价格，价格匹配流程也会另行改价。

## src_links 图片筛选

`field_rules/src_links.py` 接入 `extract_img_local.py` 的筛选逻辑，在写合并表前执行：

1. 提取 styles1 中 @ 后的图片 URL，识别转义的 #/@，避免误拆选项文本。
2. 读取 src_links 中以 # 分隔的图片，按 URL 字符串精确比较。
3. 取 src_links 独有的前 4 张，排在最前面。
4. 再追加 src_links 与 styles1 共有的全部图片，各组内部保持原始顺序。

不补入仅在 styles1 中出现的图片，不改 styles1，不下载图片。沿用原脚本的重复项语义：src_links 中重复 URL 仍保留，并占用前四张名额。styles1 为空时仅保留 src_links 前四张；src_links 为空时结果为空。

例如 src_links 为 `变体图B#普通图1#变体图A#普通图2#普通图3#普通图4#普通图5`，输出为 `普通图1#普通图2#普通图3#普通图4#变体图B#变体图A`。

日志显示来源文件、真实行号和筛选前后图片数。处理结果仅写合并表，原始脚本和来源 Excel 保持不变，不执行原脚本的固定路径或 .bak 覆盖流程。
