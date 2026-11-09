---
name: excel-a2b-converter
description: 通用 Excel「A 表 → B 表」转换技能，字段对应关系完全由用户自定义。支持 .xls/.xlsx/.csv 源表、套用目标模板保留表头与格式、按列名/列序号/Excel 列字母定位、常量填充、模板拼接（{列名}）、行号自增与 16 种数据变换。适用于「把 A 表按规则整理成 B 表」「导单模板填充」「两张表字段对齐」「大货单/销售单转换」「批量表格搬运」等场景。当用户提到 Excel 转换、表格转换、A 表转 B 表、字段映射、列对应、按模板生成表格、xls 转 xlsx、导单/对单/搬数据时使用本技能。
agent_created: true
---

# Excel A→B 转换技能

把「A 表（源表）」按**用户自定义的字段对应关系**转换写入「B 表（目标表）」。
引擎不做任何业务假设——**A→B 怎么对应，完全由配置或命令行决定**。

## 何时使用

- 用户给出两张表（A 源表、B 目标表/模板），要求"按某规则把 A 的数据整理进 B"。
- 需要套用 B 的模板（保留表头、格式），只把数据行填进去。
- 需要把一个系统的导单格式转成另一个系统的导入格式。
- 源表是 `.xls`（老格式），目标是 `.xlsx`。

## 核心概念

| 概念 | 说明 |
|------|------|
| A 表（source） | 源数据表，支持 `.xls` / `.xlsx` / `.xlsm` / `.csv` |
| B 表（target） | 目标表。给了 `template` 就套模板写数据；不给就新建工作簿 |
| header_row | 表头所在行（**1 基**）。很多表前几行是标题/元信息，务必确认 |
| mapping | 字段对应规则列表，一条 = 一个目标列 ← 一个来源 |

## 环境准备（首次执行一次）

```bash
bash <skill_dir>/scripts/bootstrap.sh
```

它会创建一个隔离 venv（默认 `~/.workbuddy/binaries/python/envs/excel-a2b`）并装好 `openpyxl`、`xlrd`。
之后统一用该 venv 的 python 运行脚本：

```bash
PY="$HOME/.workbuddy/binaries/python/envs/excel-a2b/Scripts/python.exe"   # Windows
SC="<skill_dir>/scripts/excel_a2b.py"
```

若 bootstrap 报错，退而求其次：直接用 WorkBuddy 托管解释器并手动 `pip install openpyxl xlrd`（见 `scripts/requirements.txt`）。

## 标准工作流（AI 使用）

1. **看结构**——对 A 表和 B 表分别 `inspect`，确认 sheet 名、表头行、列名。
   ```bash
   "$PY" "$SC" inspect "A.xls"
   "$PY" "$SC" inspect "B.xlsx" --json      # 结构化输出，便于取值
   ```
   > 注意：双层表头（第 1 行机器编码 + 第 2 行人工标签）时，B 表 header_row 通常取**标签行**。
2. **确认映射**——把 A 的列名、B 的列名列给用户看，**让用户明确"B 的哪列 ← A 的哪列"**。
   不要自行猜测业务对应关系；用户没说的列就不映射（保持空白）。
3. **生成 `mapping.json`**——按 `references/config-schema.md` 写。
4. **执行转换**——`"$PY" "$SC" convert --config mapping.json`。
5. **校验**——回报输出的行数、写入了哪些列（脚本会打印 `target_columns_detail`），必要时回读几行确认。

## 快速开始

```bash
# 1) 看结构
"$PY" "$SC" inspect "SH261005C-eclipse返单兰溪大货单（11月20号前出货）.xls"

# 2) 生成配置模板，人工填好后转换
"$PY" "$SC" init > mapping.json
"$PY" "$SC" convert --config mapping.json

# 3) 纯命令行（不写配置文件，人工最省事）
"$PY" "$SC" convert \
    --source "A.xls" --source-header-row 3 \
    --template "B.xlsx" --target-header-row 2 \
    --out "out.xlsx" \
    --map "目标列.数量=数量" \
    --map "目标列.序号=@index" \
    --map "目标列.标题={款式}-{货号}" \
    --const "目标列.分类=抓夹"
```

## 图形操作页（推荐人工使用）

不想敲命令时，用自带的可视化操作台：

```bash
# 方式一：本地引擎（推荐，完整保留模板格式）
python web/serve.py          # 自动打开 http://127.0.0.1:8765/

# 方式二：直接双击 web/index.html（零依赖，浏览器端转换）
```

页面能力：拖入 A 表与 B 模板 → 自动识别 sheet 与表头行 → 按同名自动匹配 →
可视化编辑映射（列 / 常量 / 模板拼接 / 表达式 / 行号 + 变换）→ 实时预览（被过滤行灰显）→ 一键转换下载。
支持导入 / 导出 `mapping.json`，与命令行引擎完全互通。

操作台进阶能力：

- **行过滤（多条件，且/或）**：步骤 3 可按列值正则跳过非数据行（如 TOTAL、备注行）；A 表含「货号」列时自动预填 `^SH\d+C-\d+$`。
- **拆行**：把某列按分隔符拆成多行（如 颜色 红/蓝 → 两行）。
- **聚合**：按某列分组、数值列求和（多行汇总为一行）。
- **表达式（expr）**：来源类型选「表达式」，如 `num(数量)*num(单价)`，中文列名可直接作变量。
- **批量转换**：A 框一次选/拖多个文件，首个为配置基准，其余套用同一配置逐个生成下载。
- **配置记忆**：转换成功后按「A 表表头指纹」自动记住配置（localStorage，最多 20 套）；下次拖入同结构文件自动套用。
- **错误提示**：源列找不到、过滤正则非法会就地标红；大文件 / 合并单元格会在日志区提示。

| 模式 | 启动方式 | 说明 |
|------|---------|------|
| 本地引擎 | `python web/serve.py` | 页面自动检测到服务后切到"本地引擎模式"，用 Python 引擎执行，**完整保留模板格式**；结果同时保存到桌面 |
| 浏览器 | 双击 `web/index.html` | 零安装；SheetJS 在浏览器内转换，能保留模板的表头内容、列宽与合并单元格，但**不保留单元格字体/边框等样式** |

> 对格式有严格要求（字体、边框、打印版式）时，用本地引擎模式或命令行引擎。

## 映射规则速查

一条规则（`mapping` 数组的元素）：

```jsonc
{ "target": "B 的目标列名", "source": "A 的源列名", "transforms": ["trim","number"], "default": "0" }
```

`source` 的几种写法：

| 写法 | 含义 |
|------|------|
| `"数量"` | 按**列名**取 A 表的值 |
| `2` / `"B"` | 按 **1 基列序号** / **Excel 列字母**取 |
| `"{款式}-{货号}"` | **模板拼接**，`{...}` 里写 A 的列名，自动替换为当前行值 |
| `"@index"` | **行号自增**（从 1 开始，`"start": 偏移` 可调） |

规则级字段：

| 字段 | 说明 |
|------|------|
| `type` | `"const"` 常量（配 `value`）/ `"template"` 拼接 / `"expr"` 表达式 / `"index"` 行号；不写则自动推断 |
| `value` | 常量值（`type=const` 时用） |
| `transforms` | 变换链，按顺序执行，见下表 |
| `default` | 取值为空时用的兜底值 |

**行加工（options / source 下，顺序：过滤 → 拆行 → 聚合）**：

- `row_filter`：按列值正则过滤，支持多条件 `{"combine":"and|or","rules":[{column,match,mode},...]}`
- `split`：`{"column":"颜色","delimiter":"/"}` —— 一行按分隔符拆多行
- `aggregate`：`{"group_by":["货号"],"sum":["数量"],"count_into":"行数"}` —— 分组求和汇总
- `type:"expr"`：`{"target":"金额","type":"expr","expr":"num(数量)*num(单价)"}` —— 行内表达式（安全沙箱，中文列名可直接作变量，`row["列名"]` 访问任意列）

详见 `references/config-schema.md`。

**变换（transforms）**：`trim` `upper` `lower` `text` `int` `number` `round:{n}` `bool`
`date:{格式}` `date_offset:{天数}` `prefix:{前缀}` `suffix:{后缀}` `replace:{[正则,替换]}` `regex_replace`
`split:{[分隔符,下标]}` `first_line` `pad:{[填充符,宽度]}`

> `transforms: ["trim","number"]` 字符串写法等价于 `[{"trim":null},{"number":null}]`，带参变换用对象写法，
> 如 `[{"date": "%Y-%m-%d"}]`、`[{"date_offset": 30}]`、`[{"round": 2}]`。

**变换预设（常用公式）**：操作台步骤 2 的「变换」列右侧带「预设 ▾」下拉，可点选去空格 / 金额两位 / 标准日期 / 补零4位 等常用公式一键套用，选中后仍可手动编辑；命令行用 `"$PY" "$SC" presets`（或 `presets --json`）查看全部内置预设清单。配置里也可直接写别名：`"transforms": "@金额两位"`，并可用 `options.transform_presets` 自定义。详见 `references/config-schema.md`「变换预设」小节。

> 双端语义已对齐：`int`/`number` 宽容取数值前缀，`round` 用 `Decimal` 四舍五入（正数 `ROUND_HALF_UP`，非 Python 内建银行家舍入；负数半值按浏览器 `Math.round` 取更接近 +∞），`first_line` 识别 `\r` / `\n` / `\r\n`。浏览器模式与本地引擎模式逐值一致。

## 关键选项（options）

| 选项 | 默认 | 说明 |
|------|------|------|
| `skip_empty_source_rows` | `true` | 跳过 A 表全空行 |
| `clear_existing_data` | `true` | 套模板时先清空模板自带的数据行（保留表头） |
| `on_missing_source` | `empty` | 源列找不到时：`empty` 留空 / `error` 报错 / `keep` 保留列名文本 |
| `create_missing_columns` | `false` | 目标表头没有该列时，自动在末尾追加 |

命令行对应开关：`--keep-empty-rows`、`--keep-template-data`、`--create-missing-columns`、`--on-missing error`。

## 常见坑

- **表头行选错**：A 表前几行常是标题/日期/工厂等元信息。用 `inspect` 看"建议表头行"，必要时 `--header-row N` 覆盖。
- **目标列名带空格/全角符号**：如 `下单明细. 产品分类`（点号后有一个空格）。必须**逐字符**与表头一致，否则匹配失败。
- **列名包含匹配**：目标列名唯一包含时才生效；有歧义时必须写全名。
- **套模板保留格式**：脚本会复制模板"首个数据行"的样式到写入的每一行。若模板首数据行是空的，样式可能偏简。
- **.xls 中的图片**：老格式图片多为 `=DISPIMG(...)` 公式文本，会被当作文本搬运，不会真的嵌入图片。

## 打包分发（给别人用）

工具已做成**单文件 exe**，对方无需装 Python / Node 即可双击使用，且完整保留模板格式。

```bash
# 1) 安装打包器（在 bootstrap 建好的 venv 里）
"$VENV/Scripts/python.exe" -m pip install pyinstaller

# 2) 构建单文件 exe（资源会一并打进 exe，无需外部文件）
cd web
"$VENV/Scripts/python.exe" -m PyInstaller --onefile --name excel-a2b --noconsole \
    --add-data "$SK/web;web" --add-data "$SK/scripts;scripts" \
    --paths "$SK/scripts" --hidden-import openpyxl --hidden-import xlrd serve.py

# 产物：dist/excel-a2b.exe
```

分发出去的内容（`excel-a2b-converter/`）：
- `excel-a2b.exe` —— 主程序，双击即用（内置引擎 + 网页，自动开浏览器）
- `web/` —— 纯网页零装版（双击 `web/index.html`，浏览器模式，样式不如 exe 版保真）
- `使用说明.txt` —— 给别人看的操作指南

> PyInstaller 冻结后，`serve.py` 通过 `sys._MEIPASS` 定位内置的 `web/` 与 `scripts/`，无需额外文件。
> 转换结果默认落在 exe 所在目录。exe 仅在用户本机 127.0.0.1 起临时服务，不上网、不上传数据。

### 跨平台（macOS / Linux）

- 当前 `excel-a2b.exe` 是 **Windows 专用**；PyInstaller 无法跨平台编译。
- 在 Mac / Linux 上想用**保真本地引擎**，无需打包：对方有 `python3` 即可运行源码
  （`bash scripts/bootstrap.sh && python3 web/serve.py`）。`serve.py` 只用标准库，跨平台。
- 想给 Mac 也做“双击程序”：在 **Mac 上**执行 `bash web/build_mac.sh`
  （注意 `--add-data` 分隔符用 `:`、`--windowed` 替代 `--noconsole`）。
- 完全零安装的纯网页版：`web/index.html` 直接用浏览器打开，任何系统都行
  （浏览器模式，保留表头/列宽/合并，不保留字体/边框样式）。

## 文件说明

| 路径 | 作用 |
|------|------|
| `scripts/excel_a2b.py` | 主引擎（inspect / convert / init / presets） |
| `scripts/bootstrap.sh` | 环境初始化 |
| `scripts/requirements.txt` | 依赖清单 |
| `web/index.html` | 可视化操作台（双击即用） |
| `web/assets/` | 操作台样式与逻辑 |
| `web/serve.py` | 本地引擎服务（保真模式） |
| `web/vendor/xlsx.full.min.js` | 离线 SheetJS（已内置，无需联网） |
| `references/config-schema.md` | 配置文件完整字段说明 |
| `examples/mapping.example.json` | 可直接改用的配置示例 |
