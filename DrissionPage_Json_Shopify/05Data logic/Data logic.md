# Styles_递增价格匹配.py 说明文档

> 核对日期：2026-09-14。以 `Data logic/Styles_递增价格匹配.py` 的实际实现为准，并用 `Data logic/styles/lovedandco_styles.xlsx` 的一批真实产物交叉验证。

## 一、这个脚本是做什么的

把采集产出的 **10 字段 styles 表**（`title / name / price1 / price2 / styles1~3 / src_links / link-href / details`）加工成可直接导入的 **Shopify CSV** 和 **WooCommerce（WP）CSV**。

与 GUI 版流水线（`03CRAWLER/Shopify_JSON_GUI.py`，详见同目录 `styles数据处理流程报告.md`）最大的区别是：

- **单文件、无 GUI、无参数**：直接改脚本里的 `INPUT_FILE` 常量后运行。
- **价格匹配发生在 Shopify 转换之前**，直接改写 styles 表本身的 `price1/price2` 列和 `styles1` 内嵌价，而不是像 GUI 那样在 Shopify 输出表上改 `Variant Price`。
- 采用 **递增分配**：把原价窗口内的价格库候选按升序依次分配给同商品的各个变体（因此文件名含“递增价格匹配”）。

脚本内 5 个步骤**顺序执行，一次跑完**。

## 二、运行方式

```bash
python "Styles_递增价格匹配.py"
```

只需修改脚本第 29 行的输入路径（唯一配置项，其余全是硬编码常量）：

```python
INPUT_FILE = r"D:\A_PythonCode\anglan_crawler\Data logic\styles\lovedandco_styles.xlsx"
```

脚本**不写 xlsx 备份/日志文件**，全部结果打印到 stdout，中间产物直接落盘。

## 三、输入与输出

### 输入

一个 10 字段的 styles xlsx，必须包含列：
`title, name, price1, price2, details, src_links, styles1, styles2, styles3`（`link-href` 可缺省，缺省时 Shopify 的 `Link-Href` 留空）。

`styles1` 的格式：

```
<Option1 Name>#<值1>&<Name2>&<值2>$price1$price2@图片#<值1'>&<Name2'>&<值2'>$price1'$price2'@图片'#...
```

- `#` 分隔组头与各变体段
- 段内 `&` 连接选项值/选项名（第 1 个值不带名称，之后是 `名称&值`）
- 段尾从右数两个 `$` 之间的两个数是内嵌价
- `@` 后是图片 URL

### 输出（全部写在**输入文件所在目录**下新开的两个子目录）

```
<input_parent>/
├── 原价/
│   ├── <stem>_原价.xlsx                  ← 第一步产物（原价修正后，10 列）
│   ├── <stem>_原价_Shopify.csv           ← 第四步产物（49 列 Shopify）
│   └── wp-<stem>_原价_Shopify.csv        ← 第五步产物（27 列 WP）
└── 价格匹配/
    ├── <stem>_价格匹配.xlsx              ← 第二步产物（价格匹配后，10 列）
    ├── <stem>_价格匹配_Shopify.csv        ← 第四步产物（49 列 Shopify）
    ├── wp-<stem>_价格匹配_Shopify.csv     ← 第五步产物（27 列 WP）
    └── <stem>_未匹配到的价格.csv          ← 第三步产物（仅 未匹配价格 一列）
```

`<stem>` 为输入文件名去掉后缀，例如输入 `lovedandco_styles.xlsx` → 输出 `lovedandco_styles_原价.xlsx`。

> 注意：**Shopify 与 WP 只输出 CSV，不额外输出 xlsx**（GUI 版是 xlsx + csv 双份）；WP 文件名是 **直接加 `wp-` 前缀**，不是 GUI 版那种 `wp-<目录>/` 子目录结构。

## 四、五个步骤详解

### 第一步：原价修正

两处修正，只做“降价价低于原价时拉平”，不做反向处理。

**(1) 行级列修正**（`price2 < price1` → `price2 = price1`）

脚本注释称 “D < C → D = C”，即 C 列 `price1`、D 列 `price2`：若 D 比 C 小，把 D 改成 C。

**(2) styles1 内嵌价修正**（`fix_styles()`）

逐个变体段做 `rsplit('$', 2)` 取两个价，若 `price2 < price1` 则把 `price2` 置为 `price1`，重建该段为 `选项$price1$price1@图片`；否则该段原样保留。返回修正的变体段数量。

统计：`C/D 修正行数` 与 `styles1 内嵌修正变体数`，随后保存到 `原价/<stem>_原价.xlsx`。

> 匹配价格时用的“原价”是 **D 列 `price2`**（见第二步），所以这一步先把 price2 抬到不低于 price1。

### 第二步：解析 + 价格匹配 + 写回

逐行（逐商品）处理 `原价.xlsx` 的 `styles1`：

**1. 解析变体**（`parse_styles()`）：返回 `(Option1 Name, [变体dict...])`，每个变体含 `option_values / price1 / price2 / image / raw`。

**2. 确定原价**：

```
original = price2（D 列）
若 price2 为空或 <= 0，回退到 price1（C 列）
```

**3. 计算每个变体应得的价格**：

| 情况 | 分配结果 |
|------|---------|
| `original` 为空或 <= 0 | 全部 `None`（该行整体跳过，不改 C/D/styles1） |
| `original > 599`（`MAX_LIB_PRICE`） | 全部变体 = `599` |
| 窗口内无候选 | 全部 `None` |
| 候选数 `n >= 变体数 V` | 取**最小的 V 个**候选 |
| 候选数 `n < V` | 按比例重复：`candidates[i*n//V]`，如 n=2、V=5 → `[c0,c0,c0,c1,c1]` |

窗口 = `original × [0.6, 1.1]`（`LOWER_RATIO=0.6`、`UPPER_RATIO=1.1`，**非对称**），候选来自升序价格库中落在窗口内的值。

**4. 写回行级 C/D 列**：

```
matched_low      = min(本行分配到的价格)  → price1（C 列）
matched_original = max(本行分配到的价格)  → price2（D 列）
```

**5. 重建 styles1**：每个命中变体改为 `选项$matched$matched_original@图片`（注意：段内 price2 统一写成**该行的最大值** `matched_original`，并非各自匹配值）；未命中的变体保留原始 `raw` 片段。

保存到 `价格匹配/<stem>_价格匹配.xlsx`。

> **与 GUI 版 `_price_match` 的关键差异**：本脚本**不做跨商品全局去重**。`used_prices` 这个集合只用于第三步统计“哪些价格没被用到”，**不会**把已用价格从候选中剔除，因此不同商品可以复用同一价格。GUI 版则是全文件去重（`global_used_prices`），且划线价 = 匹配价 × 1.2、窗口是 ±25%。

### 第三步：导出未匹配价格

```
unused = 价格库中本次未被任何商品用到的价格（升序）
```

写出 `价格匹配/<stem>_未匹配到的价格.csv`，单列 `未匹配价格`，编码 `utf-8-sig`。**只有 CSV，没有 xlsx**（与 GUI 版双份不同）。

### 第四步：转 Shopify（逻辑复制自 `03switch/01转shopify.py`）

对 `原价.xlsx` 和 `价格匹配.xlsx` **各跑一次**，分别生成 `<stem>_原价_Shopify.csv`、`<stem>_价格匹配_Shopify.csv`（49 列，已含 `Variant Image` 列）。

**主体流程**（`create_output_df()` + `convert_to_shopify()`）：

1. **Handle**：`name` 去掉非字母字符、转小写、空格换 `-` 得基名，再拼随机后缀 `-<2字母><3数字><2字母>`，本次转换内去重。
2. **字段映射**：`Title=name`、`Body (HTML)=details`、`Collection/Vendor=title`、`Variant Price=price1`、`Variant Compare At Price=price2`、`Image Src=src_links 按 # 拆分`、`Variant Weight Unit=kg`、`Published=TRUE`。
3. **Option 列**：`Option1/2/3 Name` 先取 `styles1/2/3` 的组头。
4. **笛卡尔积展开**：`Option1/2/3 Value` 做 `itertools.product`，行数 = `max(组合数, 图片数)`；多出的行只放图片。
   - 第 0 行写产品级字段（Title/Body/Collection/Vendor/Published/OptionX Name/SEO…）与固定变体字段（SKU 随机 17 位 `123AB-C12DE-456AB`、Grams=0、Tracker=shopify、Qty=999、Policy=deny、Fulfillment=manual、Requires Shipping=TRUE、Taxable=FALSE）。
   - 其余行只写固定变体字段，产品级列留空。
   - **无变体、只有图片的行**（`i >= 组合数`）会删除 SKU/Grams/库存/价格等变体列。
5. **列拆解（openpyxl 手术）**：以第 11 列 `Option1 Value`（此时存的是整段 `值1&名称2&值2$p1$p2@图`）为源，`rsplit('$',2)` 拆出价格、按 `&` 拆出选项，插入 7 列：
   `Option1 Value / Option2 Name / Option2 Value / Option3 Name / Option3 Value / Variant Price / Variant Compare At Price`，
   随后删掉原有的重复 Option 列和临时价格列，把拆出的价格再写进真正的 `Variant Price / Variant Compare At Price`。
   - 若 `Variant Compare At Price` 为空或 0 → 置为 `Variant Price`。
   - 末尾清理：`Option1 Name`（J 列）为空的行，清空 `Option2 Value`（L 列）与 `Option3 Value`（N 列）。
   - `Option1 Value` 里 `@` 之后的内容拆到 `Variant Image` 列。

> 因此本脚本**能从 styles1 段内 `&` 还原出 Option2/3 的名称与值**（如 `Option2 Name=Size`、`Option2 Value=S`），并不依赖恒为空的 `styles2/styles3`。

### 第五步：Shopify 转 WP（逻辑复制自 `03switch/03shopify转WP.py`）

对第四步的两个 Shopify CSV **各跑一次**，生成 `wp-<Shopify文件名>`（同目录），编码 `utf-8-sig`，27 列（**含 `Tags` 列**，GUI 版输出时去掉了 Tags）。

**分组**：按 `Handle` 建 `list_items`，每个 Handle 一个列表，`list_items[handle][0]` 是父级（parent），之后是各 variation。

- 父级类型：首行 `Option1 Value` 非空且非 `Default Title` → `variable`，SKU = `SKU{int(time.time())}{line_number}`；否则 `simple`，SKU = `Variant SKU`。
- `variable` 会立刻追加第一个 variation（首行也生成子变体）。
- 后续行：`Option1 Value` 非空且非 `Default Title` → 追加 variation，并把图片/选项值合并回父级（`WPAdd()` 逗号去重追加）；否则只把图片合并到父级。
- variation 名称：`父名 + "-" + (Option1,Option2,Option3 用逗号拼接)`；`Parent` 列回填父级 SKU；`Position = len(items)`（append 前长度）。
- 价格：`Sale price = Variant Price`，`Regular price = Variant Compare At Price`，为空时回退 `Variant Price`。

父级 `Attribute 1 name` 为 `Title` 时置空；`Attribute X global` 视值是否为空取 `1` 或空。

## 五、与 GUI 版流水线的差异对照

| 维度 | `Styles_递增价格匹配.py`（本脚本） | GUI 版 `Shopify_JSON_GUI.py` |
|------|-----------------------------------|------------------------------|
| 入口 | 改 `INPUT_FILE` 常量直接跑 | GUI，采集→按钮→转换 |
| 价格匹配对象 | styles 表的 `price1/price2/styles1` | Shopify 表的 `Variant Price` |
| 匹配窗口 | 原价 × **[0.6, 1.1]** | 原价 × **[0.75, 1.25]** |
| 同商品多变体 | 候选**升序分配给各变体**（递增） | 组内按价格升序，逐行取最小候选 |
| 跨商品去重 | **无**（同价可复用） | 全局去重（每价全文件用一次） |
| 行级价格 | C=min(分配给本行的价)，D=max | 逐行改，无行级聚合 |
| 划线价 | = 本行最大匹配价（D 列） | 匹配价 × 1.2 |
| 未使用价格 | 仅 CSV | xlsx + csv |
| Shopify 输出 | **仅 CSV** | xlsx + csv |
| WP 输出 | `wp-<文件名>.csv` 同目录 | `wp-<stem>/wp-<stem>.csv` 子目录 |
| WP 列 | 27 列（含 Tags） | 26 列（去掉 Tags） |
| 父 SKU | `SKU{time}{line}` | MD5 确定性哈希 `SKU+8位` |

## 六、关键常量

| 常量 | 值 | 作用 |
|------|-----|------|
| `INPUT_FILE` | 硬编码路径 | **唯一需修改的配置** |
| `price_library` | 约 215 个预定义价，8.99 ~ 599 | 匹配用价格库 |
| `LOWER_RATIO` | `0.6` | 窗口下界系数 |
| `UPPER_RATIO` | `1.1` | 窗口上界系数 |
| `MAX_LIB_PRICE` | `599` | 原价超此值时全部变体直接取 599 |

## 七、注意事项与已知坑

1. **改输入路径**：脚本无命令行参数，只认第 29 行 `INPUT_FILE`；若 `INPUT_FILE` 所在目录不含子目录名影响不大，但输出目录固定为“输入文件同级下的 `原价/`、`价格匹配/`”。
2. **无全局价格去重**：与 GUI 版相反，同一价格可被多商品复用，因此 `未匹配到的价格.csv` 的“未用”只反映本次运行，不代表库存被锁定。
3. **窗口不对称**：`[0.6, 1.1]` 意味着最低可降到原价 6 折、最高只加到 1.1 倍；原价 599 以上一律变 599。
4. **匹配价可能重复**：当窗口内候选数少于变体数时按比例重复（`candidates[i*n//V]`），同一商品可能出现重复价。
5. **行级 C/D 被聚合**：第二步后 `price1`=本行最小匹配价、`price2`=本行最大匹配价，与段内每个变体的实际价不再一一对应；真正逐变体的价在 `styles1` 段里。
6. **随机 SKU / Handle**：第四、五步每次运行都重新随机生成 SKU 与后缀，重复运行同输入会得到不同的 `Handle`、`Variant SKU`、`WP SKU`，**不能靠重复运行做增量**。
7. **原价目录也会转 Shopify/WP**：第四、五步对“原价”和“价格匹配”两份都跑，原价目录里的 Shopify/WP 是匹配前快照。
8. **第一步只降不升**：`price2 < price1` 才拉平，`price1 < price2` 不处理。
9. **原价 > 599 的写法**：走 `assigned = [599]*V` 分支，`price2` 也会被写成 599（即 max=599），即原价被“封顶”。

## 八、典型数据流示例（真实数据）

以 `lovedandco_styles.xlsx` 第 0 行（6 个变体）为例：

```
输入:
  price1(C)=102.98  price2(D)=102.05
  styles1 = Color#White&Size&S$102.98$102.05@img#...（共 6 段）

第一步: price2(102.05) < price1(102.98) → price2 = 102.98

第二步: original=102.98，窗口=[61.788, 113.278]
        候选升序 = [69, 79, 85, 89, 99, 109]（6 个，恰好 V=6）
        分配给 6 个变体 → 69 / 79 / 85 / 89 / 99 / 109
        price1 = min = 69，price2 = max = 109
        每段重建为 ...$69$109@img / ...$79$109@img / ...

第四步: Shopify 每变体 Variant Price 取段内 p1、Compare At 取段内 p2
        Option1 Value=White, Option2 Name=Size, Option2 Value=S ...
```
