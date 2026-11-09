# 配置文件（mapping.json）完整说明

所有路径可以是绝对路径，也可以是**相对于运行命令时的当前目录**。

```jsonc
{
  "source": {
    "path": "A表.xls",          // 必填：源表路径
    "sheet": "1",               // 选填：sheet 名或 0 基序号；不填取第一个
    "header_row": 3,            // 表头所在行，1 基；不填默认 1
    "data_start_row": 4,        // 数据起始行，1 基；不填默认 header_row + 1
    "end_row": null             // 只读到第几行（1 基，含）；不填读到末尾
  },

  "target": {
    "template": "B表.xlsx",     // 选填：目标模板。给了就套模板写数据；不填则新建工作簿
    "path": "输出.xlsx",        // 必填：输出文件路径
    "sheet": "Sheet1",          // 选填：目标 sheet 名；不填取第一个
    "header_row": 2,            // 目标表头行，1 基
    "data_start_row": 3         // 数据写入起始行，1 基；不填默认 header_row + 1
  },

  "mapping": [ /* 见下方「映射规则」 */ ],

  "options": { /* 见下方「选项」 */ }
}
```

> `null` 与"不写该字段"等价，都会回落到默认值。

---

## 映射规则（mapping 数组元素）

| 字段 | 类型 | 必填 | 说明 |
|------|------|:---:|------|
| `target` | string \| int | ✔ | 目标列：**列名**（与目标表头逐字符一致）／1 基**列序号**／**Excel 列字母**（如 `"G"`） |
| `source` | string \| int | — | 源数据来源（见下）。`type=const` 时不需要 |
| `type` | string | — | `const` / `template` / `expr` / `index` / `copy`。不写则自动推断 |
| `value` | any | — | `type=const` 时的常量值 |
| `start` | int | — | 仅 `@index` 用，行号偏移，默认 0（即首行 = 1） |
| `transforms` | array | — | 变换链，按顺序执行 |
| `default` | any | — | 取值为空（null/空串）时的兜底值 |

### source 写法

| 写法 | 语义 |
|------|------|
| `"数量"` | 按列名精确匹配；若无精确匹配，则尝试**唯一包含**匹配 |
| `2` | 第 2 列（1 基） |
| `"B"` | 同第 2 列（Excel 列字母） |
| `"{款式}-{货号}"` | 模板拼接；`{列名}` 替换为当前行该列的值 |
| `"@index"` | 行号自增，从 1 开始 |

> 只要 `source` 里出现 `{...}` 占位符，就自动按模板拼接处理；`{#index}`/`{index}`/`{序号}` 也是行号。
> 拼接时字面量原样保留，例如 `"SKU-{货号}"` → `SKU-A001`。

### 变换（transforms）

| 名称 | 参数 | 示例 | 结果 |
|------|------|------|------|
| `trim` / `strip` | — | `"  a "` → | `"a"` |
| `upper` / `lower` | — | `"aB"` → | `"AB"` / `"ab"` |
| `text` | — | `123` → | `"123"` |
| `int` | — | `"1,234"` → | `1234` |
| `number` | — | `"12.50"` → | `12.5`（整数则给 int） |
| `round` | 整数位数 `[n]` 或 `n` | `{"round":2}` on `3.14159` | `3.14` |
| `bool` | — | `"是"`→1 / `""`→0 | `1` |
| `date` | `[格式]`，默认 `%Y-%m-%d` | `"2026.10.5"` → | `"2026-10-05"` |
| `date_offset` | `[天数]` | `{"date_offset":30}` | 加 30 天 |
| `prefix` / `suffix` | `[文本]` | `{"prefix":"P-"}` on `"x"` | `"P-x"` |
| `replace` / `regex_replace` | `[正则, 替换]` | `{"regex_replace":["^SH-",""]}` | 去前缀 |
| `split` | `[分隔符, 下标]` | `{"split":["\n",1]}` | 取第 2 段 |
| `first_line` | — | 多行 → | 第一行 |
| `pad` | `[填充符, 宽度]` | `{"pad":["0",4]}` on `"5"` | `"0005"` |

字符串简写 `"number"` 与对象写法 `{"number": null}` 等价；带参变换必须用对象写法。

**数值解析与舍入（引擎与浏览器模式逐值对齐）**：

- `int` / `number` 对「数值后跟非法字符」采用**宽容数值前缀**解析，与浏览器 `parseInt` / `parseFloat` 一致：`"12abc"` → `12`、`"1e3"` → `1000`（`int` 对齐 `parseInt` 故得 `1`）、`"+3.5"` → `3.5`；先去掉千分位逗号，`"1,234.567"` → `1234` / `1234.567`。解析失败（`"abc"`、`""`）返回原值。
- `round` 采用 `Decimal` 四舍五入，**不是** Python 内建 `round()` 的银行家舍入：正数半值 `ROUND_HALF_UP`（`round:2` 时 `0.125` → `0.13`、`2.675` → `2.68`），负数半值按浏览器 `Math.round`「并列时取更接近 +∞」的语义（`ROUND_HALF_DOWN`，`round:2` 时 `-0.125` → `-0.12`；`round:1` 时 `-1.25` → `-1.2`）。非法输入返回原值。注意：JS 的 `round` 分支不去除千分位逗号，引擎如实复刻；需处理 `"1,234.567"` 时请串接 `number`（如预设 `@金额两位`）。
- **`n` 应为整数位数**（15 条内置预设全部为整数）。非整数 `n`（如 `{"round":2.5}`）无业务含义：浏览器 `Math.pow(10, n)` 与 Python `10**n` 在个别 `n`（如 2.5 / 3.5）上相差末位 1 ULP，两端结果可能有极微小差异，**不在支持范围**，请勿使用。
- `first_line` 识别 `\r` / `\n` / `\r\n` 三种换行（含单独 `\r`），取首行并去首尾空白。

日期解析同时支持 Excel 序列号、`%Y-%m-%d`、`%Y/%m/%d`、`%Y.%m.%d`、`%Y年%m月%d日`。

### 变换预设（常用公式）

把常用变换链做成「预设」，两种用法等价：

- **Web 操作台**：步骤 2「变换」列右侧的「预设 ▾」下拉，点选即把等价变换链写入输入框（**会替换当前变换内容**），选中后仍可继续手动编辑；最后一项「清空变换」用于一键清空。
- **配置别名**：`transforms` 直接写字符串 `"@名称"`，引擎执行时展开成等价列表。例如 `{"target":"金额","source":"单价","transforms":"@金额两位"}`。查不到的别名按**单个变换名**处理（不会逐字符拆开），因此 `"transforms":"trim"` 也是合法的。
- **命令行查阅**：`python excel_a2b.py presets` 打印清单，`presets --json` 输出机器可读版本（含 Web 等价文本）。

| 预设别名 | 等价变换链（Web 输入框文本） | 展开后的 transforms |
|---------|------------------------------|---------------------|
| `@去空格` | `trim` | `["trim"]` |
| `@转数字` | `trim,number` | `["trim","number"]` |
| `@金额两位` | `number,round:2` | `["number",{"round":2}]` |
| `@取整` | `number,int` | `["number","int"]` |
| `@转大写` | `trim,upper` | `["trim","upper"]` |
| `@转小写` | `trim,lower` | `["trim","lower"]` |
| `@转文本` | `text` | `["text"]` |
| `@标准日期` | `date:"%Y-%m-%d"` | `[{"date":"%Y-%m-%d"}]` |
| `@日期+30天` | `date_offset:30` | `[{"date_offset":30}]` |
| `@首行` | `first_line` | `["first_line"]` |
| `@补零4位` | `pad:["0",4]` | `[{"pad":["0",4]}]` |
| `@去SH前缀` | `regex_replace:["^SH",""]` | `[{"regex_replace":["^SH",""]}]` |
| `@是/否` | `bool` | `["bool"]` |
| `@加前缀P-` | `prefix:"P-"` | `[{"prefix":"P-"}]` |
| `@清空变换` | `""` | `[]` |

#### 自定义预设

可在 `options.transform_presets` 或配置顶层 `transform_presets` 里覆盖 / 追加（两处都会合并进内置预设；同名时 `options` 优先）。键可省略 `@` 前缀，值可以是变换链数组、单个变换名字符串或单个变换字典：

```jsonc
{
  "options": {
    "transform_presets": {
      "@脱敏手机号": [{"regex_replace": ["(\\d{3})\\d{4}(\\d{4})", "\\1****\\2"]}],
      "@去空格": ["trim", "text"]      // 覆盖内置同名预设
    }
  },
  "mapping": [
    {"target": "手机号", "source": "电话", "transforms": "@脱敏手机号"},
    {"target": "金额", "source": "单价", "transforms": ["trim", "@金额两位"]}
  ]
}
```

> 预设别名也可以写在 `transforms` 数组元素里（如上例 `@金额两位`）。
> 别名由 Python 引擎展开；Web 下拉写入的是等价展开链，操作台导入含**内置别名**的 mapping.json 时也会自动还原成等价链文本，
> 因此浏览器模式与本地引擎模式结果一致。自定义别名（写在 `transform_presets` 里的）仅在引擎侧展开，
> 若要在浏览器模式下使用，请直接写展开后的变换链。

---

## 选项（options）

| 字段 | 默认 | 说明 |
|------|------|------|
| `skip_empty_source_rows` | `true` | 跳过源表全空行 |
| `clear_existing_data` | `true` | 套模板时清空模板原有数据行（表头行保留） |
| `on_missing_source` | `"empty"` | 源列不存在：`empty` 留空 / `error` 抛错 / `keep` 填入列名文本 |
| `create_missing_columns` | `false` | 目标表头缺列时在末尾自动新建该列 |
| `row_filter` | `null` | 按某列格式跳过行（详见下） |
| `split` | `null` | 拆行：一行 A → 多行 B（详见下） |
| `aggregate` | `null` | 聚合：多行 A → 一行 B（详见下） |

行加工顺序：**行过滤 → 拆行 → 聚合 → 字段映射**。

### split：拆行（一行拆成多行）

把某列按分隔符拆开，每个片段生成一行（其余列原样复制）：

```jsonc
"options": {
  "split": {
    "column": "颜色",       // 拆哪一列
    "delimiter": "/",       // 分隔符，默认 "/"
    "regex": false,         // true 时 delimiter 按正则切（re.split）
    "trim": true,           // 片段去首尾空白
    "skip_empty": true      // 跳过空片段（该列整体为空的行会被丢弃）
  }
}
```

例：`颜色="红/蓝"` 的一行 → 两行，颜色分别为 `红`、`蓝`。
注意：拆行**不会**平摊数量等数值列，每行都保留原值；需要合计请配合 `aggregate`。

### aggregate：聚合（多行汇总为一行）

按分组列把多行合并成一行，指定列求和，其余列取组内第一行的值：

```jsonc
"options": {
  "aggregate": {
    "group_by": ["货号"],       // 分组列（可多个）
    "sum": ["数量", "总价"],    // 求和列（非数值按 0 计）
    "count_into": "行数"        // 选填：把组内行数写入这个（新）列名
  }
}
```

### 映射规则新类型：expr 表达式

`{"target":"金额","type":"expr","expr":"num(数量) * num(单价)"}`

- 列名若为合法标识符（**中文列名可以直接当变量**），直接在表达式里用；
- 含空格/点号的列名用 `row["下单明细. 产品分类"]` 访问；
- 可用函数：`num(x)`（转数值，失败为 0）、`text(x)`、`float/int/str/round/len/min/max/abs/sum/bool`；
- 支持条件表达式：`num(单价) * 0.95 if num(数量) > 100 else num(单价)`；
- 表达式在隔离环境求值（无内建、不能 import/读写文件）；
- 表达式出错时报错会带目标列名与行号，便于定位。

`split`/`aggregate` 也可以放在 `source` 节点下（与 `options` 等价，`options` 优先）。

### row_filter：按列值格式跳过非数据行（支持多条件）

适用于「源表末尾有几行汇总/备注（如 `TOTAL`、`客户要求`）不需要转换」的场景。
引擎逐行取 `column` 指定的列值，用 `match`（正则）判断是否匹配：

```jsonc
"options": {
  "row_filter": {
    "column": "货号",            // 用哪一列判断（必须是 A 表已存在的列名）
    "match":  "^SH\\d+C-\\d+$",  // 正则：数据行应满足的格式
    "mode":   "include"          // include：仅保留「匹配」的行；
                                 // exclude：跳过「匹配」的行
  }
}
```

- 支持多条条件：写成数组 `[{...},{...}]`，或用 `{combine:"and"|"or", rules:[...]}`；单条写法 `{column,match,mode}` 向后兼容（等价于 `and` 包一条）。`combine:"and"` 要求全部满足，`or` 任一满足即保留。
- `match` 为空串的条件会被忽略（同未配置）。正则非法会直接报错。
- 匹配采用「搜索」语义（`re.search`）：要强制整列值完全符合，请在正则里写 `^...$`。
- 统计字段新增 `skipped_by_filter`（被本条规则跳过的行数）与 `source_rows_before_filter`。

> 上面例子：货号形如 `SH261005C-1` 的是数据行，而 `TOTAL` / `客户要求` 不匹配 `^SH\d+C-\d+$`，
> 因此 `mode:"include"` 会自动只转换 9 行数据，略过末尾两行。

### 常见默认：货号列默认正则匹配

本工具对「大货单」类文件提供了一个开箱即用的默认值：当 A 表存在名为 `货号` 的列时，
操作台会自动勾选该过滤并填好 `column:"货号"`、`match:"^SH\d+C-\d+$"`、`mode:"include"`，
即**按货号默认正则保留 SH 货号行（如 `SH261005C-1`），自动略过末尾的 `TOTAL` / `客户要求` 等非数据行**。
对应到配置文件就是：

```jsonc
"row_filter": { "column": "货号", "match": "^SH\\d+C-\\d+$", "mode": "include" }
```

`^SH\d+C-\d+$` 是货号的完整格式（SH + 数字 + C + 横杠 + 序号）。`mode:"include"` 表示保留匹配行。
用户在操作台里手动改过该过滤后，默认不再自动覆盖。

---

## 命令行等价参数

| 配置文件字段 | 命令行 |
|--------------|--------|
| `source.path` | `--source` |
| `source.sheet` | `--source-sheet` |
| `source.header_row` | `--source-header-row` |
| `source.data_start_row` | `--source-data-start-row` |
| `target.template` | `--template` |
| `target.path` | `--out` |
| `target.sheet` | `--target-sheet` |
| `target.header_row` | `--target-header-row` |
| `target.data_start_row` | `--target-data-start-row` |
| `mapping[].target/source` | `--map "目标=来源"`（可重复） |
| `mapping[]{type:const}` | `--const "目标=值"`（可重复） |
| `skip_empty_source_rows=false` | `--keep-empty-rows` |
| `clear_existing_data=false` | `--keep-template-data` |
| `create_missing_columns=true` | `--create-missing-columns` |
| `on_missing_source` | `--on-missing empty\|error\|keep` |

命令行模式不支持 `transforms`（复杂变换请用 `--config`）。
