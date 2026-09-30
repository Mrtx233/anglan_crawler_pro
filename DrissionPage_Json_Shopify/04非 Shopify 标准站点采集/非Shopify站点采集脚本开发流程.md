# 非 Shopify 标准站点采集脚本开发流程

> 适用于不能直接使用标准 Shopify `{URL}.json` 的站点，例如：
>
> * Shopify Headless
> * Next.js / RSC
> * Magento PWA
> * WooCommerce
> * React / Vue SPA
> * 自研电商网站
>
> 核心目标：
>
> **只适配站点采集逻辑，最终输出主管线统一的 10 字段数据。**

---

# 一、核心原则

开发一个新站，优先解决三个问题：

```text
① 商品数据在哪里
② 变体怎么对应价格和图片
③ 页面什么时候真正加载完成
```

不要一开始就猜 XPath。

优先级：

```text
JSON / GraphQL / XHR
→ JSON-LD
→ Next.js / RSC 等内嵌数据
→ 渲染后的商品 DOM
→ 点击变体后动态获取
```

原则：

* 数据源按站点实际情况选择
* 不强制所有站点使用同一种解析方式
* 不从商品名猜颜色或尺码
* 不人为创建站点不存在的变体
* 错误页、骨架页、空数据不能当成成功商品
* 尽量只修改站点解析层，不重复开发公共功能

---

# 二、公共工具

非 Shopify 采集脚本统一使用：

```python
from non_shopify_utils import (
    clean_body_html,
    fetch_exchange_rates,
    format_image_url,
    format_price,
)
```

公共工具文件：

```text
04非 Shopify 标准站点采集/
├── non_shopify_utils.py       ← 公共工具（保留在根目录）
└── 非 Shopify 站点/           ← 各站点采集脚本
    ├── snitch_crawler_GUI.py
    ├── dobell_crawler_GUI.py
    ├── hespokestyle_crawler_GUI.py
    ├── joyfit_crawler_GUI.py
    └── xxx_crawler_GUI.py
```

采集脚本位于 `非 Shopify 站点/` 子目录，导入时把上级目录加入 sys.path：

```python
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from non_shopify_utils import (
    clean_body_html,
    fetch_exchange_rates,
    format_image_url,
    format_price,
)
```

由于脚本深了一层，路径常量也随之调整：

```python
MODULE_DIR = Path(__file__).resolve().parent   # 04非 Shopify 标准站点采集/非 Shopify 站点
SHOPIFY_ROOT = MODULE_DIR.parent.parent        # DrissionPage_Json_Shopify
DEFAULT_INPUT = str(SHOPIFY_ROOT / "01INPUT_XLSX" / "xxx.xlsx")
```

禁止继续引用：

```python
from Shopify_JSON_GUI import ...
```

禁止指向主管线的 sys.path：

```python
sys.path.insert(... "03CRAWLER")
```

`04非 Shopify 标准站点采集` 必须能够独立运行。

---

# 三、开发流程

```text
打开代表商品
↓
确认站点架构
↓
找到商品真实数据源
↓
确定 10 字段来源
↓
确认变体、价格、图片关系
↓
实现页面就绪判断
↓
实现商品解析
↓
生成 styles1
↓
完整性校验
↓
单商品测试
↓
3～5 个商品冒烟测试
↓
批量运行
```

---

# 四、先探查，不要直接写正式脚本

开发新站前，先确认下面这些数据分别在哪里：

```text
商品名
售价
划线价
颜色
尺码
其他 Option
Variant
Variant 图片
商品主图
商品详情
币种
```

常见页面特征：

| 页面特征                                  | 优先检查             |
| ------------------------------------- | ---------------- |
| `/_next/static/`、`self.__next_f.push` | Next.js / RSC    |
| `/graphql`                            | GraphQL          |
| `wp-content`、`wc-ajax`                | WooCommerce      |
| `cdn.shopify.com` 但没有 `.json`         | Shopify Headless |
| `window.Shopline`、`img.myshopline.com` | Shopline（SSR HTML） |
| 页面只有 root/app                         | XHR / Fetch      |

可以临时创建：

```text
_probe_站点.py
```

用于打印：

* HTML
* JSON-LD
* RSC / Next 数据
* XHR / GraphQL
* 商品区域 DOM
* 变体切换后的变化

探查完成后删除 `_probe` 临时脚本。

---

# 五、10 字段固定契约

最终 xlsx 必须保持以下顺序：

```text
title
name
price1
price2
styles1
styles2
styles3
src_links
link-href
details
```

即：

```python
CSV_FIELDS = [
    "title",
    "name",
    "price1",
    "price2",
    "styles1",
    "styles2",
    "styles3",
    "src_links",
    "link-href",
    "details",
]
```

---

## title

输入 xlsx 第一列原样透传。

---

## name

当前商品真实名称。

必须：

```text
非空
不是 Skeleton
不是推荐商品名称
```

---

## price1

商品实际售价，最终转换为 USD。

---

## price2

真正存在的划线价。

如果没有：

```python
price2 = price1
```

不能人为制造划线价。

---

## styles1

商品存在变体时生成。

没有变体：

```python
styles1 = ""
```

---

## styles2 / styles3

当前流程保持：

```python
""
```

多维 Option 放入 `styles1` segment 中表达。

---

## src_links

只采集当前商品主图：

```text
img1#img2#img3
```

每张图片必须经过：

```python
format_image_url()
```

---

## link-href

输入 xlsx 第二列商品 URL 原样透传。

同时作为断点和 upsert 的唯一键。

---

## details

商品详情 HTML 必须经过：

```python
clean_body_html()
```

---

# 六、价格规则

首先确认页面真实币种。

例如：

```text
INR
GBP
EUR
AUD
```

然后：

```python
rates = fetch_exchange_rates()

if not rates:
    raise RuntimeError("汇率获取失败，停止采集")
```

不要在汇率失败后继续生成 USD 数据。

---

## format_price 注意事项

当前项目保留历史 cents 规则：

```text
"999"    → 可能按 cents 处理
"999.00" → 999 主币单位
```

非标准站点 DOM 通常直接显示主币单位。

因此如果获取：

```python
price_raw = "1129"
```

应先处理：

```python
if price_raw and "." not in price_raw:
    price_raw = f"{price_raw}.00"
```

再：

```python
price1 = format_price(
    price_raw,
    rates=rates,
    currency=currency,
)
```

---

# 七、图片规则

所有输出图片必须经过：

```python
format_image_url()
```

包括：

```text
src_links 中的图片
styles1 中 @ 后面的图片
```

最终满足主管线 `_600x600` 契约。

例如：

```text
xxx.jpg?v=123
↓
xxx_600x600.jpg?v=123
```

禁止：

* 扫描整个页面所有 `<img>`
* 混入推荐商品
* 混入 Logo
* 混入支付图标
* 混入促销 Banner

优先从：

```text
商品媒体 JSON
商品 Gallery 容器
当前 Variant 数据
```

获取图片。

---

# 八、变体规则

不要假设所有商品都是：

```text
Color × Size
```

实际可能是：

```text
无变体

Size

Color

Color × Size

Style × Size

Color × Size × Length
```

必须根据目标商品真实 Option 生成。

---

# 九、styles1 格式

## 无变体

```text
styles1 = ""
```

---

## 只有 Size

```text
Size#S$10$10@img#M$10$10@img
```

---

## 只有 Color

```text
Color#Black$10$10@black#Blue$10$10@blue
```

---

## Color × Size

```text
Color#Black&Size&S$10$10@black#Black&Size&M$10$10@black#Blue&Size&S$10$10@blue
```

基本格式：

```text
option1_name
#
option1_value
&option2_name&option2_value
&option3_name&option3_value
$price1
$price2
@image
```

其中：

```text
#
&
$
@
```

属于保留分隔符。

---

# 十、是否生成笛卡尔积

如果站点数据表示 Option 可以自由组合：

```text
4 Color × 5 Size
=
20 个组合
```

则生成完整笛卡尔积。

但如果接口明确只有：

```text
17 个有效 Variant
```

则只能生成：

```text
17 个
```

不能人为补足成 20。

优先相信站点真实 Variant 数据。

---

# 十一、Variant 图片和价格

每个 segment 应尽可能使用自己的：

```text
价格
图片
```

图片优先级：

```text
Variant 图片
→ 当前 Color 图片
→ 当前商品 preview_image
→ 商品首图
```

不要简单写：

```python
first_img = images[0]
```

然后给所有颜色使用同一张图片。

---

# 十二、Snitch 多色特殊案例

Snitch RSC / Flight 数据中存在：

```text
colors
color_variants_ids
color_variants
```

例如：

```text
colors:
Blue
Khaki
Pink
White
```

正确图片映射：

```text
color
↓
color_variants_ids 同位置 product_id
↓
color_variants 中查找 shopify_product_id
↓
preview_image
```

不能：

```python
zip(colors, color_variants)
```

因为两个数组顺序可能不同。

最终应该形成：

```python
color_image_map = {
    "Blue": blue_image,
    "Khaki": khaki_image,
    "Pink": pink_image,
    "White": white_image,
}
```

然后：

```python
for color in colors:
    for size in sizes:
        ...
```

---

## Snitch 必须保留的回归案例

至少长期保留一个：

```text
Blue
Khaki
Pink
White
```

×

```text
S
M
L
XL
XXL
```

预期：

```text
4 × 5 = 20 个 styles1 segment
```

并验证：

```text
Blue → Blue 图片
Khaki → Khaki 图片
Pink → Pink 图片
White → White 图片
```

不能只验证单颜色商品。

---

## 十二·二、Shopline 特殊案例（joyfit.store）

Shopline 平台（特征：`window.Shopline`、`img.myshopline.com`、`/cdn/shop/`）的
`/products/*.json` 返回 403，但商品数据**全部内嵌在服务端渲染的 HTML** 里，
直接解析 `tab.html` 即可，无需点击变体。

### 数据源

| 字段 | 来源 |
| --- | --- |
| name / details / sku | ld+json `@type=Product`（`name` / `description` / `sku`） |
| 价格 / 币种 | ld+json `offers[0].price` + `priceCurrency`（无划线价则 price2 = price1） |
| 全部变体 | `<script name="variant-data" type="application/json">`，每条含 `options=[颜色,尺码]`、`price`（分）、`featured_media_id` |
| 颜色 / 尺码 | variant-data 的 `options[0]` / `options[1]`（按出现顺序去重） |
| 商品图 | 媒体画廊 `li[data-media-id] //img/@src`（去 query 后 `format_image_url` 去重） |

### 颜色 → 图片映射（id 桥接，不按顺序 zip）

Shopline 的变体图通过 `featured_media_id` 关联媒体画廊：

```text
variant-data: options[0]=颜色, featured_media_id=7431631181370379647
                    ↓
媒体画廊:    li[data-media-id="7431631181370379647"] //img/@src
```

形成 `color_image_map` 后，颜色 × 尺码 笛卡尔积，每段带该色图。规则同 Snitch：
不能按数组顺序直接对齐，必须用 media-id 桥接。

### 与 Snitch 的差异

| | Snitch | Shopline（joyfit） |
| --- | --- | --- |
| 数据载体 | Next.js flight 数据 | SSR HTML 内嵌 |
| 变体来源 | colors / color_variants | `<script name="variant-data">` |
| 图映射键 | shopify_product_id | featured_media_id ↔ data-media-id |
| 价格 | ld+json color/size + offers | ld+json offers（USD） |

---

# 十三、页面就绪判断

SPA / PWA 不能只判断元素是否存在。

例如：

```html
<h1>
    <span class="TextPlaceholder"></span>
</h1>
```

虽然 h1 已经存在，但商品并没有真正加载完成。

页面至少满足：

```text
name 非空
price 可解析
至少一张商品图
details 非空
当前 URL 属于当前任务商品
不是错误页
```

再开始解析。

---

## 禁止固定 sleep 判断加载完成

不建议：

```python
time.sleep(5)
```

建议：

```text
轮询业务条件
↓
成功
或
超时失败
```

---

# 十四、错误页面

发现以下页面必须直接失败：

```text
403
404
Access Denied
Cloudflare 验证
Sucuri
验证码
登录页
代理错误页
浏览器网络错误
```

不要继续解析。

---

# 十五、推荐站点脚本结构

每个新站主要只需要解决：

```python
wait_for_product()

parse_product()

extract_product()
```

---

## wait_for_product()

负责：

```text
判断页面是否正常
判断商品是否加载完成
判断 URL 是否正确
```

---

## parse_product()

负责获取原始数据：

```text
name
currency
price
compare_price
options
variants
images
details
```

尽量不要在这里处理 GUI 或 Excel。

---

## extract_product()

负责：

```text
价格转换
图片格式化
styles1 构建
details 清洗
记录完整性检查
```

最后返回统一数据结构。

---

# 十六、完整性校验

保存前至少检查：

```text
name
price1
price2
src_links
link-href
details
```

如果商品有变体：

```text
styles1
```

也必须有效。

推荐统一：

```python
is_record_complete(record)
```

只有：

```text
完整记录
```

才能加入 `done`。

---

# 十七、断点原则

```text
成功商品
→ 加入 done

失败商品
→ 可以保存 link-href 占位
→ 不加入 done

后续补采成功
→ 根据 link-href 替换失败记录
```

不能只因为输出 xlsx 已经出现 URL 就认为采集完成。

---

# 十八、测试要求

不要一开始建立很重的测试体系。

新站开发采用：

```text
1 个代表商品
↓
确认解析正确
↓
3～5 个不同商品冒烟
↓
补关键自动化回归
```

---

## 每个新站至少检查

```text
商品名
价格
币种
USD价格
Option
Variant数量
styles1 segment 数量
图片数量
详情长度
```

---

## 必测边界

根据站点实际存在的商品选择：

```text
普通商品
促销商品
无变体
单维变体
多维变体
部分缺货
不同品类
```

不需要为了测试而寻找站点不存在的商品类型。

---

# 十九、公共工具测试

`non_shopify_utils.py` 已有测试后，新增站点不需要重复测试公共函数。

公共层主要长期保证：

```text
clean_body_html
fetch_exchange_rates
format_price
format_image_url
```

正确即可。

重点测试站点自己的：

```text
页面解析
变体映射
styles1
```

---

# 二十、完成标准

新增一个非 Shopify 站点完成后，检查：

* [ ] 没有引用 `03CRAWLER`
* [ ] 没有引用 `Shopify_JSON_GUI`
* [ ] 使用 `non_shopify_utils`
* [ ] 10 字段顺序正确
* [ ] 商品名正确
* [ ] price1 / price2 为正确 USD
* [ ] 无划线价时 `price2 = price1`
* [ ] styles1 与真实 Variant 一致
* [ ] 多维 Option 组合数量正确
* [ ] 每个 Variant 图片对应正确
* [ ] src_links 只包含当前商品
* [ ] 所有图片经过 `format_image_url()`
* [ ] details 经过 `clean_body_html()`
* [ ] 错误页不会保存
* [ ] 空字段不会被当成成功
* [ ] 失败记录不会加入 done
* [ ] 至少验证 3～5 个商品
* [ ] 关键特殊 Variant 案例有回归测试

---

# 二十一、当前目录职责

```text
04非 Shopify 标准站点采集/
│
├── non_shopify_utils.py
│      └── 公共价格 / 图片 / HTML / 汇率工具（保留在根目录，供子目录脚本导入）
│
├── 非Shopify站点采集脚本开发流程.md
│
└── 非 Shopify 站点/                  ← 各站点采集脚本目录
    ├── snitch_crawler_GUI.py
    │      └── Snitch 站点解析
    │
    ├── dobell_crawler_GUI.py
    │      └── Dobell 站点解析
    │
    ├── hespokestyle_crawler_GUI.py
    │      └── He Spoke Style 站点解析（WordPress/WooCommerce 定制主题，仅尺码变体）
    │
    ├── joyfit_crawler_GUI.py
    │      └── Joyfit 站点解析（Shopline，SSR 内嵌 ld+json + variant-data + media-id 桥接）
    │
    └── xxx_crawler_GUI.py
           └── 后续新站

注：各脚本通过
`sys.path.insert(0, str(Path(__file__).resolve().parent.parent))`
引入根目录的 `non_shopify_utils`；`SHOPIFY_ROOT = MODULE_DIR.parent.parent`。
```

目前不要为了抽象而继续大规模重构。

等新增更多站点后，如果以下代码开始大量重复：

```text
Excel
GUI
浏览器
重试
断点
保存
```

再考虑抽取：

```text
crawler_base.py
```

现阶段开发新站的重点始终是：

```text
数据源
+
页面就绪
+
变体映射
+
10 字段输出
```
