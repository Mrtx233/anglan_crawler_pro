# styles 表每条数据的完整处理流程

> 核对日期：2026-09-12。以同目录 `Shopify_JSON_GUI.py` 的实际实现为准；前置链接采集章节仅作背景，本次未核对 `detail_link_GUI.py`。

## 总览

每条数据从输入到最终输出，经历 **4 个阶段**：

```
输入 Excel (title + url)
  → [阶段一] 采集 parse_product() → 10 字段 xlsx
  → [断点]   采集完成后暂停，等待用户点击「合并并转换」按钮
  → [阶段二] Shopify 转换 _styles_to_shopify() → Shopify 格式 49 列
  → [阶段三] 价格匹配 _price_match() → 价格替换后的 Shopify 格式
  → [阶段四] WP 转换 _shopify_to_wp() → WooCommerce 格式 CSV
```

### 「合并并转换」按钮（采集与转换之间的断点）

采集完成后**不再自动进入阶段二**，而是暂停等待人工确认：

```
_run_task() 采集全部完成
  → 有数据: self.pending_merge = {"folder": 文件夹, "results": 全部记录}
            启用「合并并转换」按钮，日志提示人工检查各输出文件
  → 无数据: 打印"无采集数据，跳过合并"
```

点击按钮后 `_do_merge()` 校验（采集中不可合并、必须有 pending 数据），随后在后台线程执行 `_merge_worker()`：

```
_merge_worker()
  1. save_xlsx() 保存 {输入文件夹名}_合并.xlsx 到 02OUTPUT_XLSX 对应目录
  2. _styles_to_shopify(合并表) → Shopify 转换、价格匹配、尝试 WP 转换
  3. 完成后清空 pending_merge，状态栏提示"合并并转换完成"
  未被内部捕获的异常: 打印异常堆栈，重新启用按钮允许重试
  WP 转换异常在 _styles_to_shopify 内部捕获，仅记日志，外层仍可能提示完成
```

**实际合并来源**：`pending_merge["results"]` 是采集结束时的内存记录，按钮不会重新读取已输出的拆分 xlsx。暂停后直接修改拆分表不会进入本次合并；需要重新启动采集，使旧输出重新载入。重跑默认跳过已完成 URL，修改图片 KEEP/SKIP 也不会重新加工旧记录。

手动停止后，只要已有 `all_results`，也可合并当前部分结果；尚未处理文件的旧输出不会自动补入。开始新任务会清空待合并状态。合并直接拼接各文件记录，不按 title 或 URL 做跨文件去重。

---

## 前置步骤：详情页链接采集（detail_link_GUI.py）

流水线的输入 Excel（`title` + `url`）由 `detail_link_GUI.py` 从 Shopify 列表页采集生成。它与 `Shopify_JSON_GUI.py` 独立运行，只产出链接表，不参与 styles 加工。

### 输入
GUI 文本框中的分类列表，每行 `标题, 列表页URL`，如 `Loungewear > Pajamas, https://example.com/collections/pajamas`。

### 采集方式
每个分类逐页抓取（`?page=N`，可用「最大页数」限制），单页链接来自三种来源去重合并：
1. 页面 `var meta = {...}` JSON 的 `products[].handle`
2. HTML 全文正则 `"handle":"..."`（捕获懒加载追加的数据）
3. DOM `/products/xxx` 链接（捕获 JS 渲染的卡片，含 Shadow DOM）

填了 XPath 则用 XPath 提取；否则滚动加载后用 handle 提取。无限滚动分类持续触发滚动/点 Load More 直到链接数稳定。

### 输出
按域名分组，每个域名一个 `{域名}.xlsx`，表头 `title / link` 两列，供 `read_targets_from_excel()`（按列位置读 col0=title、col1=url）直接消费。

### 保存与断点（合并去重，不覆盖）
保存前 `load_existing_domain_rows()` 读已存在的同名文件（按 link 去重成 `{link: title}` 映射），与本次结果合并后再写回：
- 已有 link 保留原 title，不重复追加
- 新增 link 追加到末尾
- 日志 `已保存: xxx（已有 X 条 + 新增 Y 条 = 合并 Z 条）`

因此重复运行只追加新链接，不会覆盖历史采集结果；文件不存在或读取失败则重新生成，不阻断流程。

---

## 阶段一：采集（parse_product）

### 入口
`_run_task()` 读取活动工作表，从第二行开始按列位置取 col0=title、col1=url，去两端空白并跳过空 URL。`group_targets_by_json_url()` 去掉请求 URL 的查询参数、fragment 和尾部斜杠，补 `.json` 后在单个输入文件内去重，同一请求地址保留首次 title。随后通过 `fetch_json()` 获取 JSON，调用 `parse_product(data, table_title, original_url)` 生成一条记录。

已有输出通过 `link-href` 与分组后的 `original_url` 精确比较实现续采。每累计 20 次成功触发保存，文件处理结束也保存已有结果；最终成功/失败数仅统计本次请求，合并条数包含已载入的历史记录。输入、输出路径通过字符串替换 `01INPUT_XLSX → 02OUTPUT_XLSX` 得到，所选路径需要包含该目录名，否则输出可能与输入同路径。

### 产物：10 字段

| 字段 | 来源 | 处理逻辑 |
|------|------|---------|
| `title` | 输入 Excel 的 `title` 列 | 读取/分组时去两端空白，parse_product 内原样传入 |
| `name` | `product["title"]` | Shopify JSON 中产品的标题 |
| `price1` | fallback variant 的 `price` | `format_price()` → 汇率转换（非 USD 时除以汇率）→ 无小数点时除以 100 视为分 → 去尾零 |
| `price2` | fallback variant 的 `compare_at_price` | `format_price2()` → 优先取 `compare_at_price`（非零时），否则回退到 `price`，同样经过 `format_price()` |
| `styles1` | `build_variant_combo()` | **详见下方** |
| `styles2` | 硬编码 `""` | 始终为空 |
| `styles3` | 硬编码 `""` | 始终为空 |
| `src_links` | `build_first_image_srcs()` + `build_images_by_color_option()` | `#` 分隔的图片 URL 列表 |
| `link-href` | 输入 URL 去 query/fragment | 原始商品页面链接 |
| `details` | `product["body_html"]` | `clean_body_html()` → 去除 `<a>` 标签但保留链接文字，移除 `<img>`；其他标签仅保留标签名、丢弃全部属性；文本 HTML 转义，移除码点 > U+FFFF 的字符，循环清理空标签 |

### styles1 的构建过程（build_variant_combo）

**步骤 1 — get_sorted_options(product)**

```
输入: product["options"]
处理:
  1. 按 position 排序
  2. 根据传入的 skip_options 过滤名称包含关键词的 option（不区分大小写）；采集入口传 SKIP_OPTIONS
  3. 只保留 values 数量 > 1 的 option（即多值选项）
  4. 尺码名称统一: 如果 option 名称含 "size" 或 values 全部是尺码值(S/M/L/XL等)，将 name 统一为 "Size"
  5. Type 改名: 如果 option 名称为 "Type"（脚本中的导入兼容处理），统一改为 "Style"
  6. 把名为 "color" 的 option 移到最前面
输出: multi_value options 列表
```

### 尺码统一规则（`_is_size_option`）

两个条件满足其一即视为尺码选项：

| 条件 | 说明 | 示例 |
|------|------|------|
| 名称含 "size" | 不区分大小写 | `Size`、`SIZE`、`Product Size`、`xxx Size` |
| values 全是尺码值 | 所有值在 SIZE_VALUES 集合中 | values 为 `["S", "M", "L"]`，即使名称是 "Choose" 也会被识别 |

尺码值集合（`SIZE_VALUES`）：
```
XS, S, M, L, XL, XXL, XXXL, XXXXL, 2XL, 3XL, 4XL, 5XL,
XS/S, S/M, M/L, L/XL, One Size, OS, Free Size
```

被识别为尺码后，`option["name"]` 被统一改为 `"Size"`，确保最终输出的 styles1 中尺码名称一致。

### Type → Style 改名规则

当前脚本为规避 Type 属性导入问题，在 `get_sorted_options` 中把名为 `"Type"`（不区分大小写）的多值 option 统一改名为 `"Style"`：

- **尺码判断优先级更高**：名为 `Type` 但 values 全是尺码值（如 `S/M/L`）时，先按尺码规则统一为 `Size`，不会再改成 `Style`
- 改名后 styles1 前缀相应变为 `Style#...`

**步骤 2 — build_option_combos(options)**

```
输入: 过滤后的 options
处理: 对所有 option 的 values 做笛卡尔积
示例: options = [Color: [Red, Blue], Size: [S, M, L]]
      → 6 个 combo: (Red,S), (Red,M), (Red,L), (Blue,S), (Blue,M), (Blue,L)
输出: combo 元组列表
```

这里展开全部笛卡尔积，不根据实际 variants 列表或库存筛掉不存在的组合；缺失组合仍按后续规则回退取价。选项重排保留原 position，匹配时按原 `option{position}` 取值。

**步骤 3 — build_images_by_color_option(product, options)**

```
输入: product + options
处理:
  1. 主选项 = 名为 "color" 的 option；没有 color 时退回 options[0]（第一个多值选项）
     （get_sorted_options 已把 color 排到第 0 位，故 options[0] 永远是主选项）
  2. build_images_by_id(product) → 过滤图片（KEEP/SKIP），将图片 id→url 映射
  3. 遍历 variants，对每个 variant 取主选项对应字段（option{position}）的值
  4. 如果该 variant 有 image_id 且在 images_by_id 中存在，记录 {选项值: image_url} 映射
输出: {选项值: image_url} 字典；同一主选项值只保留首个命中的图片，其他选项组合共用该图
说明: 没有 color 选项时也照常建映射（如 Device/Size 单选项商品），
     让每个 variant 按自己的 image_id 带上对应图；映射键是主选项值，
     build_segment 用 combo[0] 取值命中
```

**步骤 4 — 遍历 combo 构建 segment**

```
对每个 combo:
  1. find_variant_by_combo() → 精确匹配所有 option 字段
  2. 找不到 → find_variant_by_first_option() → 只匹配排序后首个保留选项对应的 option{position}（不一定是原始 option1）
  3. 还找不到 → 使用 fallback_variant（优先 position=1，否则首个 variant，无 variants 则为空字典）
  4. build_segment() 生成一个 segment 字符串
```

**步骤 5 — build_segment 格式**

```
格式: option1_value&option2_name&option2_value$price1$price2@image_url

示例:
  Red&Size&S$19.99$24.99@https://cdn.shopify.com/..._600x600.jpg

其中:
  - option1_value 不带名称前缀
  - option2 开始带名称: name&value
  - 价格用 $ 分隔
  - 图片 URL 用 @ 后缀
```

> **转义规则（选项值/选项名含分隔符时）**
> 选项值或选项名若含保留字符 `&` `#` `@` `\`（如 `Tie & Square`），
> 写入时会在该字符前加反斜杠转义：`Tie & Square` → `Tie \& Square`；
> 组头选项名（第 0 位）同样转义。价格与图片 URL 不转义；`$` 不在保留字符集合中，价格仍靠最右两个 `$` 解析。
> 解析时只按未转义的分隔符切分，再把 `\&` `\#` `\@` `\\` 还原为原字符，
> 因此含 `&` 的选项值不再被误切乱码。旧文件记录若不含保留字符，解析结果与改动前完全一致。

**步骤 6 — 最终拼接**

```
如果 option1 有名称:
  "Color#Red&Size&S$19.99$24.99@img1.jpg#Blue&Size&M$19.99$24.99@img2.jpg"
否则:
  "Red&Size&S$19.99$24.99@img1.jpg#Blue&Size&M$19.99$24.99@img2.jpg"
```

**无多值选项的情况**：`build_option_combos()` 返回 `[()]`，styles1 通常只有 `$price1$price2`。转换时它被当作单独组头，随后清空 `$` 开头的名称，变体行价格回退到商品行的 `price1/price2`，不会生成主选项图片映射。

### src_links 的构建过程

两条路径合并：

**路径 A — build_first_image_srcs(product)**
```
1. 所有图片按 position 排序
2. _filter_images() 过滤（KEEP/SKIP）
3. 如果有 variant_id 参数且匹配到图片，再筛选；没匹配到则保留原列表（采集入口未传该参数）
4. 每个图片 URL 调用 format_image_url() → 在扩展名前插入 _600x600，保留 query/fragment；无扩展名不插入，已有尺寸后缀也不做去重
5. 用 # 拼接
```

**路径 B — build_images_by_color_option(product, options)**
```
1. 同 styles1 的步骤 3，按主选项（color 或首个多值选项）得到 {选项值: image_url} 映射
2. 没有多值选项时返回空字典
3. 取所有 values 去重后加入 src_links
```

**最终 src_links** = 路径 A 的结果 + 路径 B 中不在路径 A 中的 URL

### 图片过滤（_filter_images）

```
KEEP 非 None → 只保留 position 在 KEEP 列表中的图片（空列表 [] 会过滤全部）
KEEP 为 None、SKIP 非 None → 排除 position 在 SKIP 列表中的图片
两者均为 None → 不过滤；KEEP 始终优先
负数 → 按排序去重后的 position 值倒数定位；越界忽略
GUI 留空得到 None；非空但未解析出整数的 KEEP 可能得到 []
```

### 价格处理（format_price）

```
1. 仅 rates 和 variant.price_currency 都有值时尝试换汇；非 USD 且有有效汇率则除以汇率并先保留两位小数
2. 如果包含 "." → 用 float 解析
3. 如果不包含 "." → 视为"分"，除以 100（如 "1999" → 19.99）
4. 最多保留两位小数并去尾零: "19.90" → "19.9", "20.00" → "20"
5. 缺少币种/汇率则保持原币金额进入上述格式化；无法解析的文本原样返回
```

---

## 阶段二：Shopify 转换（_styles_to_shopify）

### 入口
采集完成后暂停，用户点击「合并并转换」按钮 → `_merge_worker()` 保存合并 xlsx，然后调用 `_styles_to_shopify(merge_path, print)`。

### 输入：阶段一的 10 字段 xlsx

函数校验其中 9 列；`link-href` 可缺省，届时 Shopify 的 `Link-Href` 留空。

### 整体结构

`_styles_to_shopify` 是阶段二的主函数，它内部依次调用阶段三和阶段四：
```
_styles_to_shopify(input_file)
  ├── 1. 读取 10 字段 xlsx
  ├── 2. 逐行展开为 Shopify 49 列表格
  ├── 3. 保存 _Shopify_原价.xlsx / .csv
  ├── 4. 调用 _price_match() → 价格匹配           ← 阶段三
  ├── 5. 调用 _shopify_to_wp(原价CSV) → WP 转换    ← 阶段四
  ├── 6. 查找价格匹配CSV，存在才调用 WP 转换（当前路径存在偏差，见下文）
  └── 返回 原价 xlsx 路径
```

### 处理流程（逐行详细）

对输入 xlsx 的每一行（一个 product），执行以下步骤：

---

**步骤 1 — 生成 Handle**

```
_clean_handle_name(name)
  输入: item["name"]，如 "Women's Summer Dress"
  处理: re.sub(r"[^a-zA-Z\s]", "", name).lower() → 去掉数字和特殊字符，转小写
        .replace(" ", "-").strip("-") → 空格换 "-"
  输出: "womens-summer-dress"

_generate_unique_handle(base, used_handles)
  输入: "womens-summer-dress"
  处理: 随机后缀按 2字母+3数字+2字母拼接 → "womens-summer-dress-ab123cd"
  去重: used_handles 集合保证本次转换内唯一
  输出: "womens-summer-dress-ab123cd"
```

---

**步骤 2 — 解析 styles1/styles2/styles3**

`_parse_style_group(value)` 对每个 styles 字段执行：

```
输入: "Color#Red&Size&S$19.99$24.99@img1.jpg#Blue&Size&M$19.99$24.99@img2.jpg"

处理:
  1. _clean_cell(value) → 去两端空白
  2. _style_split(value, "#") → ["Color", "Red&Size&S$19.99$24.99@img1.jpg", "Blue&Size&M$19.99$24.99@img2.jpg"]
  3. 去掉空白项

如果 parts 为空 → 返回 ("", [""])
如果 parts 只有 1 个 → 返回 (parts[0], [""])  ← 只有名称，没有值
如果 parts >= 2   → 返回 (parts[0], parts[1:]) ← 名称 + 值列表

结果:
  option1_name = "Color"
  option1_values = ["Red&Size&S$19.99$24.99@img1.jpg", "Blue&Size&M$19.99$24.99@img2.jpg"]
```

对 styles2/styles3 同理，但通常为空字符串，返回 `("", [""])`。

**特殊规则**: 如果 option_name 以 `"$"` 开头（如 `"$19.99"` 被误解析为名称），则清空为 `""`。

> **解析转义**: 步骤 2 与步骤 5c 的 `split` 都只切在未转义的分隔符上；
> 段内选项值与选项名最终会还原反斜杠转义（见步骤 5 的转义规则），
> 保证 `Tie \& Square` 这类值还原为 `Tie & Square`。

---

**步骤 3 — 收集图片**

```
images = src_links.split("#") → ["img1_600x600.jpg", "img2_600x600.jpg", "img3_600x600.jpg"]
去空白项
```

---

**步骤 4 — 笛卡尔积展开**

```
combinations = itertools_product(option1_values, option2_values, option3_values)

option1_values = ["Red&Size&S$19.99$24.99@img1.jpg", "Blue&Size&M$19.99$24.99@img2.jpg"]
option2_values = [""]
option3_values = [""]

→ combinations = [
    ("Red&Size&S$19.99$24.99@img1.jpg", "", ""),
    ("Blue&Size&M$19.99$24.99@img2.jpg", "", ""),
  ]

total_rows = max(len(combinations)=2, len(images)=3, 1) = 3
```

**关键**: `total_rows` 取 combinations 和 images 的较大值。如果图片比 variant 多，多出的行只有图片没有 variant 数据。

---

**步骤 5 — 逐行填充 Shopify 49 列**

对 `index` 从 0 到 `total_rows-1`（本例 index=0,1,2），每行初始化 49 列为空字符串。

**5a. 每行都填的字段：**

| 字段 | 值 | 说明 |
|------|-----|------|
| `Link-Href` | 原始商品 URL | 每行相同 |
| `Handle` | `"womens-summer-dress-ab123cd"` | 每行相同 |
| `Variant Weight Unit` | `"kg"` | 写死 |

**5b. 仅 index=0 填的字段（产品级信息，只出现在第一行）：**

| 字段 | 来源 |
|------|------|
| `Title` | `item["name"]`，如 `"Summer Dress"` |
| `Body (HTML)` | `item["details"]`，如 `"<p>Beautiful summer dress</p>"` |
| `Collection` | `item["title"]`，如 `"lindvs.com连衣裙"` |
| `Vendor` | `item["title"]`，同上 |
| `Published` | `"TRUE"` |
| `Gift Card` | `"FALSE"` |
| `SEO Title` | = Title |
| `SEO Description` | = Title |

**5c. 如果 index < len(combinations) — 有 variant 数据的行：**

取 `combo = combinations[index]`，对 `combo[0]`（第一个 segment）调用 `_parse_variant_segment` 解析：

```
_parse_variant_segment("Red&Size&S$19.99$24.99@img1.jpg")

步骤:
  1. _clean_cell → 去空白
  2. 按第一个未转义的 "@" 拆分: text="Red&Size&S$19.99$24.99", image="img1.jpg"
  3. 按 "$" 从右拆两次: text="Red&Size&S", price1="19.99", price2="24.99"
     (rsplit("$", 2) 从右边开始切，最多切 2 次)
  4. 用 _style_split 按未转义的 "&" 拆分 text，再逐字段还原转义: ["Red", "Size", "S"]

返回:
  {
    option1_value: "Red",
    option2_name:  "Size",
    option2_value: "S",
    option3_name:  "",
    option3_value: "",
    price1: "19.99",
    price2: "24.99",
    image: "img1.jpg",
  }
```

然后填入字段，每个字段都有回退逻辑：

| 字段 | 优先值 | 回退值 |
|------|--------|--------|
| `Option1 Name` | option1_name（"Color"） | — |
| `Option1 Value` | segment 解析的 option1_value | combo[0] 原值 |
| `Option2 Name` | segment 解析的 option2_name | option2_name（styles2 的名称） |
| `Option2 Value` | segment 解析的 option2_value | combo[1] 原值 |
| `Option3 Name` | segment 解析的 option3_name | option3_name（styles3 的名称） |
| `Option3 Value` | segment 解析的 option3_value | combo[2] 原值 |
| `Variant Price` | segment 解析的 price1 | item["price1"]（行级回退） |
| `Variant Compare At Price` | segment 解析的 price2 | item["price2"]（行级回退） |
| `Variant Image` | segment 解析的 image | — |

固定值字段：

| 字段 | 值 |
|------|-----|
| `Variant SKU` | `_generate_variant_sku()` — 随机 17 字符，格式 `123AB-C12DE-456AB` |
| `Variant Grams` | `0` |
| `Variant Inventory Tracker` | `"shopify"` |
| `Variant Inventory Qty` | `999` |
| `Variant Inventory Policy` | `"deny"` |
| `Variant Fulfillment Service` | `"manual"` |
| `Variant Requires Shipping` | `"TRUE"` |
| `Variant Taxable` | `"FALSE"` |

**5d. 如果 index < len(images) — 有图片的行：**

| 字段 | 值 |
|------|-----|
| `Image Src` | `images[index]` |
| `Image Position` | `index + 1` |

**重要**: variant 和 image 是独立判断的。一行可以同时有 variant 数据和图片，也可以只有其中之一。

---

### SKU 生成详解

#### Variant SKU（`_generate_variant_sku`）

阶段二 `_styles_to_shopify` 中为每个 variant 行随机生成。格式 `123AB-C12DE-456AB`（17 字符，含两个连字符）：

```
┌──────┬──────┬──────┐
│ 123AB │ C12DE │ 456AB │
└──────┴──────┴──────┘

生成规则:
  部分1: 1位非零数字(1-9) + 2位数字(00-99) + 2位大写字母(AA-ZZ)
  部分2: 1位大写字母(A-Z) + 2位数字(00-99) + 2位大写字母(AA-ZZ)
  部分3: 3位数字(000-999) + 2位大写字母(复用部分1的字母)

  例: 部分1 letters1="AB" → 部分3 也使用 "AB"
```

**特点**：纯随机，每次运行生成不同的 SKU。本次 `_styles_to_shopify()` 的 `used_skus` 集合去重，冲突时自动重新生成。

#### Parent SKU（`_generate_parent_sku`）

阶段四 `_shopify_to_wp` 中为 variable 类型的 parent 行生成。格式 `SKU12345678`（固定 11 位）：

```
生成规则:
  1. hash_input = f"{handle}_{counter}"  (如 "summer-dress-abc123_0")
  2. MD5 哈希 → 32 位十六进制字符串
  3. 提取所有数字字符
  4. 不足 8 位 → 末尾补 "0" 至 8 位
  5. 取前 8 位 → "SKU" + digits[:8]
```

**特点**：相同 handle、counter 得到相同候选 SKU；本次 WP 转换内冲突时递增 counter。重新执行 styles 转换会随机生成新 handle，父 SKU 也可能变化。simple 行优先复用 Variant SKU，缺失时仍生成父 SKU。

#### 对比

| | Variant SKU | Parent SKU |
|---|---|---|
| 生成阶段 | 阶段二 `_styles_to_shopify` | 阶段四 `_shopify_to_wp` |
| 格式 | `123AB-C12DE-456AB` | `SKU12345678` |
| 算法 | 随机拼接 | MD5 确定性哈希 |
| 唯一性范围 | 单次转换内去重（跨产品） | 单次转换内去重（跨产品） |
| 确定性 | 否，每次运行不同 | 同 handle、counter 相同；冲突时 counter 递增 |
| 用途 | 每个 variant 行 | variable 父级行 |
| simple 类型 | 直接使用 Variant SKU | Variant SKU 缺失时用作回退 |

---

**步骤 6 — 输出**

```
保存路径:
  {input_path.parent}/xlsx/{input_stem}_Shopify_原价.xlsx
  {input_path.parent}/csv/{input_stem}_Shopify_原价.csv

输出列: 49 列 SHOPIFY_COLUMNS
```

### 完整示例：一行 product → 多行 Shopify

```
输入行:
  title:  "lindvs.com连衣裙"
  name:   "Summer Dress"
  price1: "19.99"
  price2: "24.99"
  styles1: "Color#Red&Size&S$19.99$24.99@img1.jpg#Blue&Size&M$19.99$24.99@img2.jpg"
  styles2: ""
  styles3: ""
  src_links: "img1.jpg#img2.jpg#img3.jpg"
  details: "<p>Beautiful summer dress</p>"

↓ 解析

option1_name="Color", option1_values=["Red&Size&S$19.99$24.99@img1.jpg", "Blue&Size&M$19.99$24.99@img2.jpg"]
option2_name="", option2_values=[""]
option3_name="", option3_values=[""]
images=["img1.jpg", "img2.jpg", "img3.jpg"]
combinations=[("Red&Size&S$19.99$24.99@img1.jpg","",""), ("Blue&Size&M$19.99$24.99@img2.jpg","","")]
total_rows = max(2, 3, 1) = 3

↓ 输出 3 行

Row 0 (index=0):
  Title="Summer Dress", Body(HTML)="<p>Beautiful...</p>", Collection="lindvs.com连衣裙"
  Option1 Name="Color", Option1 Value="Red"
  Option2 Name="Size", Option2 Value="S"
  Variant Price="19.99", Variant Compare At Price="24.99"
  Variant Image="img1.jpg", Variant SKU="123AB-C12DE-456AB"
  Image Src="img1.jpg", Image Position=1

Row 1 (index=1):
  (Title/Body/Collection 为空 — 非第 0 行)
  Option1 Name="Color", Option1 Value="Blue"
  Option2 Name="Size", Option2 Value="M"
  Variant Price="19.99", Variant Compare At Price="24.99"
  Variant Image="img2.jpg", Variant SKU="789FG-H34IJ-012FG"
  Image Src="img2.jpg", Image Position=2

Row 2 (index=2):
  (无 variant 数据 — index=2 >= len(combinations)=2)
  (Title/Body 等为空)
  (Option/Variant 等为空)
  Image Src="img3.jpg", Image Position=3
  ← 只含图片的纯图片行
```

---

## 阶段三：价格匹配（_price_match）

### 入口
`_styles_to_shopify()` 内部调用 `_price_match(str(output_path), log)`，传入阶段二生成的 `_Shopify_原价.xlsx`。

### 目的
把 Shopify 产品原有的真实价格，替换为 `PRICE_LIBRARY` 价格库中预定义的标准化价格，同时生成划线价（Compare At Price）。

### 处理流程（逐步）

---

**步骤 1 — 读取并准备数据**

```
输入: _Shopify_原价.xlsx
读取: _read_table() → pandas DataFrame
校验: 必须包含 "Handle" 和 "Variant Price" 列
```

**步骤 2 — 确定输出文件名前缀**

```
从文件名中去掉后缀:
  "_Shopify_待匹配" / "_Shopify_原价_拆分" / "_Shopify_原价" / "_原价"
  → 按上述顺序，只移除第一个匹配的尾缀，得到 source_stem
```

**步骤 3 — 类型转换**

```
df["Variant Price"] → object 类型（允许混合小数和空值）
df["Variant Compare At Price"] → object 类型（如果不存在则创建）
df["匹配价格"] → 新增列，全为 None，记录匹配结果供核对
```

**步骤 4 — 按 Handle 分组**

```
df.groupby("Handle")
  同一 Handle = 同一个产品，其所有 variant 行归为一组
  跳过 Handle 为 NaN 的行
```

**步骤 5 — 对每组执行价格匹配**

对每个 Handle 分组，执行以下子步骤：

**5a. 收集有效价格行**
```
遍历组内每一行:
  price = pd.to_numeric(row["Variant Price"], errors="coerce")
  如果 price 能转为数字 → 加入 valid_rows 列表 [(行索引, 价格数值)]
  不能转换的行（空值、非数字文本）→ 跳过，不参与匹配
```

**5b. 准备可用价格库**
```
available_prices = PRICE_LIBRARY 中所有未被 global_used_prices 使用的价格
（全局去重：同一价格值在整个文件中只使用一次）
```

**5c. 按价格升序逐个匹配**
```
valid_rows 按价格从小到大排序:
  对每个 (idx, original_price):
    1. 计算匹配区间: [original_price × 0.75, original_price × 1.25]
    2. 在 available_prices 中筛选落在区间内的候选
    3. 取 candidates[0]（最小的候选）作为 matched
    4. 如果 candidates 为空 → 跳过该行，不做匹配
```

**5d. 更新价格**
```
如果匹配成功:
  global_used_prices.add(matched)          ← 全局标记已使用
  available_prices.remove(matched)         ← 从可用列表中移除

  df.at[idx, "匹配价格"] = matched                                      ← 记录匹配结果
  df.at[idx, "Variant Price"] = matched                                ← 售价 = 匹配价格
  df.at[idx, "Variant Compare At Price"] = round(matched × 1.2, 2)     ← 划线原价 = 匹配价 × 1.2
```

**关键设计点：**

- **全局去重**: `global_used_prices` 是跨 Handle 的全局集合，确保整个文件中每个价格值最多被使用一次
- **分组顺序**: pandas 默认按 Handle 排序遍历，并非原始商品行顺序；随机 Handle 可能影响跨商品价格分配
- **组内排序**: 同 Handle 内按价格从低到高匹配，低价 variant 优先获取低价候选
- **匹配区间 0.75~1.25**: 原始价格 ±25% 范围内找匹配，差距太大则跳过
- **取最小候选**: `candidates[0]` 取排序后最小的，不是最近价格；候选全部高于原价时仍会涨价

### 示例：价格匹配全过程

```
PRICE_LIBRARY(简化为): [9.99, 12.99, 14.99, 19.99, 24.99, 29.99]

Handle "summer-dress-abc123" 有 3 个 variant:
  Row 0: Variant Price = 20.00 (原始价格)
  Row 1: Variant Price = 15.00
  Row 2: Variant Price = 25.00

valid_rows = [(Row0, 20.00), (Row1, 15.00), (Row2, 25.00)]
按价格排序: [(Row1, 15.00), (Row0, 20.00), (Row2, 25.00)]

匹配 Row1 (15.00):
  区间: [11.25, 18.75]
  候选: [12.99, 14.99]
  → matched = 12.99
  Variant Price = 12.99, Compare At Price = 15.59
  global_used_prices = {12.99}

匹配 Row0 (20.00):
  区间: [15.00, 25.00]
  候选: [19.99, 24.99]（14.99 小于下限 15.00）
  → matched = 19.99
  Variant Price = 19.99, Compare At Price = 23.99
  global_used_prices = {12.99, 19.99}

匹配 Row2 (25.00):
  区间: [18.75, 31.25]
  候选: [24.99, 29.99]（19.99 已被占用）
  → matched = 24.99
  Variant Price = 24.99, Compare At Price = 29.99
  global_used_prices = {12.99, 19.99, 24.99}
```

### 产物

```
输出路径:
  {base_dir}/xlsx/{source_stem}_Shopify_价格匹配.xlsx
  {base_dir}/csv/{source_stem}_Shopify_价格匹配.csv
  {base_dir}/xlsx/{source_stem}_Shopify_未匹配到的价格.xlsx  ← 未被使用的价格
  {base_dir}/csv/{source_stem}_Shopify_未匹配到的价格.csv
```

### PRICE_LIBRARY（价格库）

硬编码的 215 个预定义价格，范围 8.99 ~ 599：

| 区间 | 示例值 |
|------|--------|
| 8.99 ~ 50.81 | 密集分布，间距不固定（8.99, 9.99, 10.99, 11.99, 12.99, ...） |
| 55 ~ 99 | 稀疏分布（55, 59, 69, 79, 85, 89, 99） |
| 109 ~ 599 | 大额价格（109, 119, 129, 189, 219, 259, 299, 319, 349, 369, 399, 459, 499, 519, 599） |

---

## 阶段四：WP 转换（_shopify_to_wp）

### 入口
`_styles_to_shopify()` 先转换原价 CSV，再尝试转换价格匹配 CSV：

1. `_shopify_to_wp(原价CSV)` → 对全部匹配前记录做 WP 转换。
2. 价格匹配路径存在才调用第二次。**当前代码只把返回的 `xlsx/...价格匹配.xlsx` 改为 `xlsx/...价格匹配.csv`，但实际 CSV 保存于 `csv/`，正常新输出目录下第二次转换会被跳过。** 手动调用 `_shopify_to_wp()` 并传入正确的 `csv/...价格匹配.csv` 才可生成对应 WP 文件；此处记录现状，未修改脚本。

WP 异常被内部捕获，且路径不存在时静默跳过，因此日志“Shopify 转 WP 完成”不代表两个版本都已生成。

### 目的
把 Shopify 格式（一行一个 variant + 图片）转换为 WooCommerce 格式（Parent 行 + Variation 子行）。

### 处理流程（逐步）

---

**步骤 1 — 读取 Shopify 数据**

```
_read_shopify_rows(input_file)
  xlsx → pd.read_excel → fillna("") → 每行转为 dict
  csv  → csv.DictReader → 每行转为 dict
返回: [{col: value, ...}, ...]
```

**步骤 2 — 按 Handle 分组，构建 products 字典**

```
products = {}

遍历每一行 row:
  handle = row["Handle"]
  Option1 Value = row["Option1 Value"]

  空 handle → 跳过
  如果 handle 不在 products 中 → 创建新的 product 条目:
    products[handle] = [parent]  ← 列表，第一个元素是 parent

  无论是否首次出现，都累积图片；Option1 Value 非空且不为 Default Title 时追加 variation
```

**步骤 3 — 创建 Parent 行**

对每个新产品（handle 首次出现），判断类型并创建 parent：

**3a. 判断 variable vs simple**

```
is_variable = (Option1 Value != "" 且 Option1 Value != "Default Title")

variable: 首行有非默认 Option1 Value；不根据变体数量判断
simple:   首行 Option1 Value 为空或 Default Title；类型由首行决定
```

**3b. 生成 SKU**

```
variable → _generate_parent_sku(handle)
  基于 handle 的 MD5 hash 取前 8 位数字 → "SKU12345678"
  同一 handle、counter 返回相同候选 SKU；冲突时递增 counter

simple → 直接用 row["Variant SKU"]（或生成 parent SKU 作为回退）
```

**3c. 价格处理**

```
variable 的 parent:
  Sale price = ""     ← parent 不设价格，由子 variation 各自定价
  Regular price = ""

simple 的 parent:
  Sale price = Variant Price
  Regular price = Variant Compare At Price（或回退到 Variant Price）
```

**3d. Option 名称处理**

```
raw_opt1_name = row["Option1 Name"]

如果 raw_opt1_name == "Title" 或以 "$" 开头 → 清空为 ""
（"Title" 是 Shopify 的默认 option 名，无实际意义；
 "$" 开头表示被误解析的价格字符串）

否则 → 保留原值
```

**3e. Parent 完整字段映射：**

| WP 字段 | 值 | 说明 |
|---------|-----|------|
| `Type` | `"variable"` 或 `"simple"` | |
| `SKU` | `"SKU12345678"` | deterministic hash |
| `Name` | `row["Title"]` | |
| `Published` | `"1"` | |
| `Visibility in catalog` | `"visible"` | |
| `Description` | `row["Body (HTML)"]` | |
| `In stock?` | `"1"` | |
| `Stock` | `"99999"` | |
| `Sale price` | `""` (variable) 或 `Variant Price` (simple) | |
| `Regular price` | `""` (variable) 或 `Compare At Price` (simple) | |
| `Categories` | `row["Collection"]` 或 `row["Type"]` | |
| `Tags` | `row["Tags"]` | |
| `Images` | `row["Image Src"]` | 后续会累积合并 |
| `Parent` | `""` | parent 行没有父级 |
| `Position` | `"0"` | |
| `Attribute 1 name` | opt1_name（已处理 "Title"/"$" 的情况） | |
| `Attribute 1 value(s)` | `""` (Default Title) 或 `Option1 Value` | |
| `Attribute 1 visible` | `"1"` 如果 name 非空 | |
| `Attribute 1 global` | `"1"` 如果 name 非空 | |
| `Attribute 2/3 name` | `Option2/3 Name` | |
| `Attribute 2/3 value(s)` | `Option2/3 Value` | |
| `Attribute 2/3 visible` | `"1"` 如果 name 非空 | |
| `Attribute 2/3 global` | `"1"` 如果 name 非空 | |

---

**步骤 4 — 累积合并 Parent 的图片和属性值**

每处理一行，如果该行有图片或属性值，回写到 parent：

```
parent["Images"] = _csv_join_unique(parent["Images"], row["Image Src"])
  → 逗号分隔，去重追加

如果 Option1 Value 非空且不是 "Default Title":
  parent["Attribute 1 value(s)"] = _csv_join_unique(parent["Attribute 1 value(s)"], Option1 Value)
  parent["Attribute 2 value(s)"] = _csv_join_unique(parent["Attribute 2 value(s)"], Option2 Value)
  parent["Attribute 3 value(s)"] = _csv_join_unique(parent["Attribute 3 value(s)"], Option3 Value)
```

这样 parent 最终汇集了所有 variant 的图片和所有属性值。

---

**步骤 5 — 创建 Variation 行**

对每个 `Option1 Value` 非空且不为 `Default Title` 的输入行创建 variation，**包括首次创建 parent 的那一行**。纯图片行只累积 parent 图片，不创建子行。

**5a. Variation 名称**

```
_wp_title(row) 生成名称后缀:
  取 Option1 Value, Option2 Value, Option3 Value
  过滤空值
  用 "," 拼接: "Red,S" 或 "Blue,M"
  最终: Parent Name + "-" + 后缀
  例: "Summer Dress-Red,S"
```

**5b. Variation 价格**

```
Sale price = Variant Price
Regular price = Variant Compare At Price（或回退到 Variant Price）
```

**5c. Variation 完整字段映射：**

| WP 字段 | 值 |
|---------|-----|
| `Type` | `"variation"` |
| `SKU` | `row["Variant SKU"]` |
| `Name` | `parent["Name"] + "-Red,S"` |
| `Published` | `"1"` |
| `Visibility in catalog` | `"visible"` |
| `Description` | `""` |
| `In stock?` | `"1"` |
| `Stock` | `"99999"` |
| `Sale price` | `row["Variant Price"]` |
| `Regular price` | `row["Variant Compare At Price"]` 或 `row["Variant Price"]` |
| `Categories` | parent["Categories"] |
| `Tags` | `""` |
| `Images` | `row["Variant Image"]` |
| `Parent` | `parent["SKU"]` |
| `Position` | 在 products[handle] 列表中的序号 |
| `Attribute 1/2/3 name` | 复制 parent 的值 |
| `Attribute 1/2/3 value(s)` | 该 variant 的具体值 |
| `Attribute 1/2/3 visible` | `""` |
| `Attribute 1/2/3 global` | `"1"` 如果 value 非空 |

---

**步骤 6 — 价格校验**

```
遍历所有 products 的所有行（parent + variations）:
  sale = float(Sale price)
  regular = float(Regular price)
  如果 sale > 0 且 regular > 0 且 sale > regular:
    Regular price = Sale price  ← 修正不合逻辑的定价
    price_fix_count += 1
```

---

**步骤 7 — 输出**

```
输出路径:
  {input_path.parent}/wp-{base_stem}/wp-{input_stem}.csv

输出列: WP_HEADERS（去掉 "Tags" 列）
  26 列: Type, SKU, Name, Published, Visibility in catalog, Description,
          In stock?, Stock, Sale price, Regular price, Categories, Images,
          Parent, Position, Attribute 1 name~global, Attribute 2 name~global,
          Attribute 3 name~global

输出顺序: 按 products 字典顺序，每个 product 先 parent 后 variations
```

### 完整示例：Shopify → WP

SKU 为格式示意；输入两条变体行分别带对应 Variant SKU。

```
输入 Shopify (3 行，同一 Handle):
  Row 0: Title="Summer Dress", Handle="summer-dress-abc123",
         Option1 Name="Color", Option1 Value="Red",
         Option2 Name="Size", Option2 Value="S",
         Variant Price=12.99, Variant Compare At Price=15.59,
         Variant Image="img1.jpg", Image Src="img1.jpg"
  Row 1: Option1 Value="Blue", Option2 Value="M",
         Variant Price=14.99, Variant Compare At Price=17.99,
         Variant Image="img2.jpg", Image Src="img2.jpg"
  Row 2: Image Src="img3.jpg" (纯图片行)

↓ 转换

Parent (Row 0 首次出现):
  Type=variable, SKU=SKU12345678, Name="Summer Dress",
  Sale price="", Regular price="",
  Images="img1.jpg,img2.jpg,img3.jpg"  ← 累积合并
  Attribute 1 name=Color, value(s)="Red,Blue"
  Attribute 2 name=Size, value(s)="S,M"

Variation 1 (Row 0):
  Type=variation, SKU=123AB-C12DE-456AB,
  Name="Summer Dress-Red,S",
  Sale price=12.99, Regular price=15.59,
  Images="img1.jpg", Parent=SKU12345678, Position=1

Variation 2 (Row 1):
  Type=variation, SKU=789FG-H34IJ-012FG,
  Name="Summer Dress-Blue,M",
  Sale price=14.99, Regular price=17.99,
  Images="img2.jpg", Parent=SKU12345678, Position=2

(Row 2 是纯图片行，Option1 Value 为空，不会创建 variation — 仅贡献图片到 parent)
```

### 价格校验示例

```
如果某 variation:
  Sale price = 19.99
  Regular price = 15.99  ← 划线价竟然比售价还低，不合理

价格校验:
  sale(19.99) > regular(15.99) → Regular price = 19.99
  price_fix_count += 1
```

---

## 完整数据流示例

以下使用人为简化为两个 segment 的 styles 中间数据串联演示；若源 JSON 同时有 Color=[Red, Blue] 和 Size=[S, M]，采集实际会展开四个组合。父 SKU 使用占位示意值，不代表该 Handle 的真实哈希结果。

```
输入 Excel:
  title="lindvs.com连衣裙", url="https://lindvs.com/products/summer-dress"

↓ 阶段一

styles 10 字段:
  title: "lindvs.com连衣裙"
  name: "Summer Dress"
  price1: "19.99"
  price2: "24.99"
  styles1: "Color#Red&Size&S$19.99$24.99@img1.jpg#Blue&Size&M$19.99$24.99@img2.jpg"
  styles2: ""
  styles3: ""
  src_links: "img1.jpg#img2.jpg#img3.jpg"
  link-href: "https://lindvs.com/products/summer-dress"
  details: "<p>Beautiful summer dress</p>"

↓ 阶段二

Shopify 格式 (多行):
  Handle: "summer-dress-xk137ab"
  Row 0: Title="Summer Dress", Option1=Color/Red, Option2=Size/S, Variant Price=19.99, Image Src=img1.jpg
  Row 1: Option1=Color/Blue, Option2=Size/M, Variant Price=19.99, Image Src=img2.jpg
  Row 2: Image Src=img3.jpg (无 variant，仅图片)

↓ 阶段三

价格匹配:
  假设只有该商品且价格库尚未被占用，区间为 [14.9925, 24.9875]
  Row 0: 19.99 → 15.31，划线价 = 18.37
  Row 1: 19.99 → 15.52，划线价 = 18.62（15.31 已被占用）

↓ 阶段四

WP 格式（手动向 _shopify_to_wp 传入正确的价格匹配 CSV 路径）:
  Parent: Type=variable, SKU=SKU12345678, Name="Summer Dress", Images="img1.jpg,img2.jpg,img3.jpg"
  Variation 1: Name="Summer Dress-Red,S", Sale price=15.31, Regular price=18.37, Parent=SKU12345678
  Variation 2: Name="Summer Dress-Blue,M", Sale price=15.52, Regular price=18.62, Parent=SKU12345678
```

---

## 关键全局变量

| 变量 | 作用 | 设置位置 |
|------|------|---------|
| `SKIP_POSITIONS` | 跳过图片的 position 列表 | GUI 输入框 → `_run_task()` 临时设置 |
| `KEEP_POSITIONS` | 保留图片的 position 列表 | GUI 输入框 → `_run_task()` 临时设置 |
| `SKIP_OPTIONS` | 跳过 option 的关键词列表 | 代码中硬编码为 `["ships from"]`（屏蔽"发货地"选项），未接入 GUI |
| `_exchange_rates_cache` | 汇率缓存 | `fetch_exchange_rates()` 在任务开始时调用 |

---

## 最终产物结构

正常自动链路会生成原价 WP；价格匹配 WP 受上述路径问题影响，需要传入正确路径另行转换。下图为包含该额外转换结果的完整产物示意。输出目录 `02OUTPUT_XLSX/{批次}/{操作员}/{站点名}/` 下会生成一套完整产物。以 `A0812/周晓东/alvaaros.com女装-裙子/` 为例：

```
alvaaros.com女装-裙子/
├── alvaaros.com女装-裙子_合并.xlsx        ← ① 最终汇总表（10 列）
├── everlyandcoboutique.com.xlsx          ← ② 各原始来源站拆分表
├── lunaandsage.com.au.xlsx
├── xlsx/                                 ← ③ 两份商品表 + 未使用价格清单（Excel 版）
│   ├── alvaaros.com女装-裙子_合并_Shopify_原价.xlsx
│   ├── alvaaros.com女装-裙子_合并_Shopify_价格匹配.xlsx
│   └── alvaaros.com女装-裙子_合并_Shopify_未匹配到的价格.xlsx
└── csv/                                  ← ④ 对应 CSV 版 + WP 导入文件
    ├── alvaaros.com女装-裙子_合并_Shopify_原价.csv
    ├── alvaaros.com女装-裙子_合并_Shopify_价格匹配.csv
    ├── alvaaros.com女装-裙子_合并_Shopify_未匹配到的价格.csv
    └── wp-alvaaros.com女装-裙子_合并/      ← ⑤ WooCommerce 导入目录
        ├── wp-alvaaros.com女装-裙子_合并_Shopify_价格匹配.csv  ← 需传入正确 CSV 路径另行转换
        └── wp-alvaaros.com女装-裙子_合并_Shopify_原价.csv
```

### ① 合并表（`_合并.xlsx`）

表头：`title / name / price1 / price2 / styles1 / styles2 / styles3 / src_links / link-href / details`（10 列）。

这是把本次已处理文件的历史记录和新增抓取记录直接拼接的总表，不按 `title` 去重，也不做跨文件 URL 去重。各列含义：

| 列 | 含义 |
|----|------|
| `title` | 商品分类（如 `Sales > Dresses Sale`） |
| `name` | 商品名 |
| `price1` / `price2` | 原始价格字段（匹配前的价格） |
| `styles1~3` | 变体展开，如 `Size#Small$98$98#Medium$98$98#Large$98$98` |
| `src_links` | 图片 URL（`#` 分隔） |
| `link-href` | 商品原链接 |
| `details` | 商品描述 HTML |

### ② 原始来源站拆分表

每个来源站独立一个 xlsx（`everlyandcoboutique.com.xlsx`、`lunaandsage.com.au.xlsx`），是合并前各站点抓取的原生数据，用于回溯溯源。

### ③④ Shopify 输出文件（xlsx + csv 双份）

这些文件不是按匹配成功/失败拆分的三份商品子集：

| 文件 | 含义 |
|------|------|
| `_Shopify_原价` | 全部商品和图片行的匹配前快照（49 列） |
| `_Shopify_价格匹配` | 保留全部行；命中行替换售价和划线价，未命中行保留原值；新增 `匹配价格` 列，共 50 列 |
| `_Shopify_未匹配到的价格` | 价格库中未被使用的数值清单，只有 `未匹配价格` 一列；不是失败商品表，仅有剩余价格时生成 |

### ⑤ WooCommerce 导入 CSV（`wp-*/` 目录）

最终要导入 WordPress/WooCommerce 商店的文件，采用 WooCommerce 标准导入格式（26 列：`Type / SKU / Name / ... / Attribute 1 name / Attribute 1 value(s) / ...`）。

variable 商品展开为「Parent + Variations」，simple 商品仅输出一条商品行；首个变体也会生成 variation：

| 行类型 | `Type` | 特征 |
|--------|--------|------|
| `variable`（父体） | `variable` | Name=商品名；`SKU`=MD5 确定性生成的父 SKU（如 `SKU64431415`）；`Parent` 列为空；`Attribute 1 name=Size`，value 列为全部变体值 `Small,Medium,Large` |
| `variation`（变体） | `variation` | Name 带变体后缀（如 `...-Small`）；`Parent` 列回填父 SKU；`Attribute 1 value(s)` 为该变体单个值 `Small` |

示例（每商品 1 父 + N 变体）：

```
variable   | SKU64431415 | EVERY LITTLE DETAIL DRESS   | Parent: (空)        | Size | Small,Medium,Large
variation  | 945LF-F19KH-123LF | EVERY LITTLE DETAIL DRESS-Small | Parent: SKU64431415 | Size | Small
variation  | 946NB-E96RH-456NB | EVERY LITTLE DETAIL DRESS-Medium | Parent: SKU64431415 | Size | Medium
variation  | 500KJ-V61SU-789KJ | EVERY LITTLE DETAIL DRESS-Large | Parent: SKU64431415 | Size | Large
```

导入 WooCommerce 后即为带 SKU、图片、描述、分类、变体的完整可售商品。

> 交付给上架环节的最终文件是 ⑤ 的 WP CSV；①~④ 是中间产物/备份。