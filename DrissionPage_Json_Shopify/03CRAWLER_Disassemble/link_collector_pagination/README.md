# Shopify 列表页链接采集器

`link_collector_pagination` 是一个基于 Tkinter 与 DrissionPage 的桌面工具：按分类列表逐页访问 Shopify 列表页，提取商品详情页链接，去重后按域名保存为 Excel，并维护分类来源索引。

本项目负责“分类页 → 商品链接”。商品名称、价格、变体、图片和描述由配套的 `shopify_scraper_workbench` 采集。两个工具通过 Excel 文件衔接，不会自动相互启动。

本文以当前目录中的 Python 源码为准。项目由原单文件版本拆分而来，但本文重点说明当前功能、操作方式和实际限制。

## 1. 功能与适用范围

- 一次输入多个分类，可包含多个站点，按输入顺序串行采集。
- 使用 `page` 查询参数分页，可限制每个分类的最大访问页数。
- 支持整页 HTML 提取与自定义 XPath 两种模式。
- 对页内、跨页、跨分类商品链接去重；保存时与已有域名文件合并。
- 保留同域名、同标题下的多个分类 URL。
- 支持停止请求、分类进度、实时日志和损坏 Excel 的 recovery 另存。
- 采集核心可脱离 GUI，在 Python 脚本中调用。

适用于能够通过分页 URL 访问、商品路径为 `/products/<handle>` 的页面。当前不包含自动滚动、点击“加载更多”、游标分页、登录流程或验证码处理，也不是通用站点链接爬虫。

## 2. 环境与启动

需要带 Tkinter 的 Python、可被 DrissionPage 启动的 Chromium 系浏览器，以及 `DrissionPage`、`openpyxl`。当前目录未提供依赖锁定文件，不能据此保证某个最低 Python 或依赖版本组合。

在本项目目录中执行：

```powershell
python -m pip install DrissionPage openpyxl
python main.py
```

可用以下命令检查 Tkinter；检查窗口关闭后再启动程序：

```powershell
python -m tkinter
```

当前机器的项目目录为：

```text
D:\A_PythonCode\anglan_crawler_pro\DrissionPage_Json_Shopify\03CRAWLER_Disassemble\link_collector_pagination
```

默认浏览器代理是 `http://127.0.0.1:7897`。使用前在 `config.py` 中核对 `PROXY_SERVER`；不需要显式浏览器代理时设为 `None` 或 `""`，程序便不传入 `--proxy-server`。修改配置后重新启动工具。

## 3. 快速使用

1. 运行 `python main.py`，打开“详情页链接采集”。
2. 填写保存路径；留空会使用 `DrissionPage_Json_Shopify/output`，并回填到输入框。
3. 粘贴分类列表，每行一个“标题, URL”。
4. 最大页数留空表示不限制；也可输入正整数，例如 `3`。GUI 不接受 `0`、负数或非整数。
5. XPath 留空使用默认提取；需要限定商品区域时填写 XPath。
6. 点击“开始采集”，观察分类进度和运行日志。
7. 等待保存结束，检查输出文件。中途结束时点击“停止”，等待界面恢复就绪并确认保存日志，再关闭窗口。

建议先用一个分类、最大页数 `2` 验证站点适配和输出内容，再扩大任务。

### 分类输入格式

```text
# 注释行不会采集
Women > Dresses, https://example.com/collections/dresses
Women > Tops, https://example.com/collections/tops?sort_by=best-selling
https://another.example/collections/new-arrivals
```

解析规则由 `collector.parse_categories()` 实现：

- 空行和以 `#` 开头的行跳过；找不到 HTTP/HTTPS URL 的行也跳过。
- 每行取第一个匹配的 URL，之前的文本清理首尾括号、引号、逗号后作为标题。
- 没有标题时，使用 URL 路径最后一段；路径以 `/` 结尾时可能得到空标题，建议显式填写标题。
- 这不是完整的 CSV 解析器；URL 中的空格、逗号、引号等会影响提取。
- `Women > Dresses` 只是原样保存的分类文本，不会被解析成分类树。
- 同名分类允许重复输入，但 Excel 的 `title` 无法区分同标题的不同来源页；需要区分时使用不同标题。

### XPath 模式

可使用返回链接元素或 `href` 属性节点的表达式，例如：

```xpath
//a[starts-with(@href, '/products/')]/@href
```

表达式在浏览器中通过 `document.evaluate()` 求值。链接元素取 `href`，属性节点取值，其他节点尝试取文本。结果仍需符合 `/products/<handle>` 路径规则；填写 XPath 不会取消 Shopify 商品路径过滤。

XPath 语法错误会记录为当前分类异常，并继续后续分类。表达式返回数字、布尔值等非节点集合时不适用于当前实现。

## 4. 分页、提取和去重规则

### 分页流程

每次任务创建一个浏览器，复用同一标签页，依次处理所有分类：

```text
分类输入 → 访问分页 → 等待页面稳定 → 提取链接
         → 当前分类去重 → 全任务去重 → 保存到内存
         → 下一页 / 下一分类 → 关闭浏览器 → 合并写文件
```

- 逻辑第 1 页原样使用输入 URL；从第 2 页起设置或替换 `page=N`，保留其他查询参数和 fragment。
- 若输入 URL 已带 `page=5`，第 1 次访问仍是该 URL，下一次会改为 `page=2`；因此应输入分类首页，而不是中间页。
- 首页无链接时额外重试最多 3 次，每次等待 2 秒。
- 后续页只要没有“该分类尚未见过”的链接，就立即结束该分类；不等待连续多个空页。
- 达到每分类最大页数、收到停止请求或分类发生异常时，也会结束相应处理。
- 页面上下文丢失时尝试重新加载一次；仍无法恢复则结束当前分类。

### 默认 HTML 提取

读取 `document.documentElement.outerHTML`，按以下顺序补充商品 handle：

1. HTML 中 `href` 指向 `/products/...` 的链接。
2. 符合代码匹配格式的 `var meta` 对象中的 `products[].handle`。
3. 其他 JSON 片段中的 `handle`，但要求页面同时出现对应的精确商品路径。

最终生成 `当前页面协议://当前页面域名/products/<handle>`，去掉商品链接查询参数和 fragment。默认模式与 XPath 模式均基于当前页面域名归一化，不用于保留外站商品链接。

默认扫描整个页面，导航、推荐区等位置的商品链接也可能被采集；可用更精确的 XPath 限定范围。仅出现 `/collections/.../products/...` 等不同路径结构时，当前规则可能漏采。

### 去重与分类归属

| 范围 | 规则 |
|---|---|
| 页内、分类内 | 按规范化商品 URL 去重 |
| 本次任务跨分类 | 同一链接只保留首次采集到的标题 |
| 与历史 Excel 合并 | 已有链接优先，保留旧标题，只追加新链接 |
| 分类索引 | 同域名、同标题下按提交顺序追加未记录的分类 URL |

**这不是分页断点续采。** 重跑仍会从输入首页重新请求，旧 Excel 只在任务收尾保存时读入合并。修改标题后重跑也不会自动改写旧链接的标题。

## 5. 输出文件与保存语义

```text
所选保存目录/
├── example.com.xlsx
├── another.example.xlsx
├── links_index.json
└── example.com.recovery-YYYYMMDD-HHMMSS.xlsx  # 旧表无法读取时
```

### 商品链接 Excel

每个域名一个工作簿，首行为固定表头，后续每行一个商品链接：

| title | link |
|---|---|
| Women > Dresses | https://example.com/products/sample-dress |

只写 `title`、`link` 两列。读取旧文件使用活动工作表，从第 2 行按前两列读取，不按表头名匹配。旧表中的额外列、样式和其他工作表不会在重写时保留。

### 分类索引

`links_index.json` 为 UTF-8 JSON，结构如下：

```json
{
  "example.com": {
    "Women > Dresses": [
      "https://example.com/collections/dresses",
      "https://example.com/collections/evening-dresses"
    ]
  }
}
```

索引记录本次提交的全部分类，包括没有采到链接或尚未处理的分类，不能作为“采集成功清单”。旧版字符串值会兼容为数组；旧索引无法读取时会记录日志并重建，不会为 JSON 创建 recovery 备份。

### 保存时机与 recovery

- 每个完整分页先提交到内存，**不会每页写盘**。
- 正常结束、通过停止按钮退出循环或捕获运行异常后，在 `finally` 中尝试保存此前完成分页。
- 停止时当前正在采集的页会被放弃。
- 旧域名 Excel 读取失败时，保护原文件，将本轮数据另存为带时间戳的 recovery 文件；重名时追加序号。
- 写入失败记录在日志和 `SaveReport.error` 中。`CollectorRunner.run()` 返回了行数据不代表写盘成功，需检查保存日志与文件。

当前 Excel 和 JSON 均直接写目标文件，尚未采用临时文件加原子替换。关窗处理虽发送停止请求，但随即销毁窗口，没有等待 daemon 采集线程结束；直接关窗或强制结束进程不能保证最终保存完成。

## 6. 配置参考

配置集中在 `config.py`：

| 配置项 | 默认值 | 作用 |
|---|---|---|
| `PROXY_SERVER` | `http://127.0.0.1:7897` | 浏览器代理 |
| `PORT_RANGE` | `(9222, 9322)` | 随机调试端口范围 |
| `BROWSER_START_RETRIES` | `3` | 浏览器启动尝试次数 |
| `DISABLE_IMAGES` | `True` | 禁止浏览器加载图片 |
| `USE_INCOGNITO` | `True` | 隐身模式 |
| `MAX_RETRIES` / `RETRY_DELAY` | `3` / `2.0` 秒 | 首页空结果重试 |
| `PAGE_SETTLE_WAIT` | `0.35` 秒 | 读取页面前的稳定等待 |
| `CONTEXT_LOST_RETRY_DELAY` | `0.8` 秒 | 页面上下文恢复等待 |
| `JS_DEFAULT_RETRIES` | `2` | JS 初次调用之外的重试次数 |
| `JS_CONTEXT_LOST_DELAY` / `JS_OTHER_ERROR_DELAY` | `0.6` / `0.3` 秒 | JS 重试等待 |
| `LINK_INDEX_FILENAME` | `links_index.json` | 分类索引文件名 |
| `DEFAULT_OUTPUT_DIRNAME` | `output` | GUI 默认输出目录名 |
| `XLSX_HEADERS` | `["title", "link"]` | 输出表头 |

`CollectConfig` 保存单次任务的 `categories`、`output_dir`、`max_pages`、`xpath`、`proxy`、`port_range` 和 `disable_images`。代码调用时 `max_pages=0` 表示不限制；GUI 中对应留空。

## 7. 模块结构与 Python 调用

| 文件 | 职责 |
|---|---|
| `main.py` | `LinkCollectorApp` 界面、输入校验、线程启动、停止和日志展示 |
| `config.py` | 常量、`CategoryItem`、`CollectConfig` |
| `collector.py` | 分类解析、分页构造、链接提取、`CollectorRunner` 主循环 |
| `browser_utils.py` | 创建浏览器、中断加载、JS 重试和 XPath 求值 |
| `file_utils.py` | 索引维护、旧表读取、合并保存、recovery、`SaveReport` |
| `logger.py` | 日志 sink 注册、广播、stdout 重定向与恢复 |

`main → collector → browser_utils/file_utils`；公共配置和日志由各层使用。采集核心不导入 Tkinter，文件层不依赖浏览器。

以下代码在本项目目录运行，或者保存为同目录中的独立脚本：

```python
from config import CategoryItem, CollectConfig
from collector import CollectorRunner

cfg = CollectConfig(
    categories=[CategoryItem("Dresses", "https://example.com/collections/dresses")],
    output_dir=r"D:\data\01INPUT_XLSX\demo",
    max_pages=2,
    xpath="",
    proxy=None,
)
runner = CollectorRunner(
    cfg,
    on_progress=lambda index, total, title: print(index, total, title),
)
rows = runner.run()  # 阻塞执行；内部会尝试保存
print(f"本次采集 {len(rows)} 条")
```

实际使用时将示例域名换成目标站点。其他线程可调用 `runner.interrupt()` 请求停止。未注册日志 sink 时，日志回退到标准输出；默认不会生成日志文件。

两个项目均有名为 `config`、`logger` 等的顶层模块，宜分别作为独立进程运行，避免直接混入同一解释器导致模块名冲突。

## 8. 与商品工作台衔接

1. 将采集器保存目录设为例如 `D:\data\01INPUT_XLSX\demo`。
2. 确认其中的域名 Excel 前两列为分类标题和商品 URL。
3. 在 `shopify_scraper_workbench` 中选择此文件夹，读取文件并开始采集。
4. 工作台会忽略 JSON 索引，只读取该目录第一层的 `.xlsx`。
5. 若目录存在 recovery 或其他非输入 Excel，应先核对并移出待处理目录，避免作为新的输入重复采集。

## 9. 常见问题与验证范围

| 现象 | 检查方向 |
|---|---|
| 浏览器启动失败 | DrissionPage 安装、浏览器可用性、调试端口和浏览器启动日志 |
| 浏览器能启动但页面失败 | `config.py` 代理地址、目标站点是否可访问 |
| 首页始终没有链接 | 页面是否为拦截页、路径是否符合 `/products/`、HTML 是否加载完成、XPath 是否匹配 |
| 分类提前结束 | 后续页没有新链接就结束；检查分页参数、渲染等待和站点是否只支持滚动加载 |
| 同一商品只属于一个分类 | 全任务去重保留第一次出现的分类；历史文件中的分类又优先于本次结果 |
| 已显示就绪但找不到结果 | 查看保存路径和保存失败日志；没有商品时可能只生成索引 |
| 出现 recovery 文件 | 旧 Excel 无法读取；检查后人工合并，原文件未被覆盖 |
| 保存失败或文件占用 | 关闭正在打开结果文件的 Excel，检查路径权限和磁盘空间 |

当前目录没有随附 `tests/`、依赖锁定文件或可直接运行的测试套件。旧文档中的测试数量与“全部通过”结论不作为当前版本的验证依据。实机验收可检查：两页采集、重复链接合并、停止后的保存，以及损坏旧表时另存 recovery。测试请使用单独输出目录；本文不声明已完成在线站点采集验证。
