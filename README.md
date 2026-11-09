# Excel A→B Converter

**配置驱动的通用 Excel 表格转换工具**：把 A 表（源表）按**你自己定义的字段对应关系**转换写入 B 表（目标模板）。

不猜业务、不写死映射——A 的哪一列对应 B 的哪一列，完全由配置、命令行或可视化操作台决定。

![Python](https://img.shields.io/badge/Python-3.9%2B-blue) ![License](https://img.shields.io/badge/license-MIT-green) ![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)

---

## 为什么需要它

日常把「大货单 / 导单 / 对账单」整理成另一个系统的导入模板时，经常要：改列名、拼接字段、加常量、算金额、跳过末尾的合计与备注行、保留 B 模板的表头与格式。手工复制粘贴既慢又容易错。

这个工具把这些动作变成一份可复用的 `mapping.json`，一次配置、反复使用；也可以完全不装 Python，直接在浏览器里拖文件转换。

## 特性

| 类别 | 能力 |
|---|---|
| 源表格式 | `.xls` / `.xlsx` / `.xlsm` / `.csv`；目标输出 `.xlsx` |
| 列定位 | 列名（精确 / 唯一包含匹配）、1 基列序号、Excel 列字母（`G`） |
| 来源类型 | 取列、常量、模板拼接（`{款式}-{货号}`）、行号 `@index`、**表达式 `expr`** |
| 数据变换 | 16 种变换 + **15 条常用公式预设**（操作台下拉点选即用） |
| 行加工 | 多条件正则过滤（且 / 或）、**拆行**（一行→多行）、**聚合**（多行→一行求和） |
| 目标写入 | 可套用 B 模板（表头、字体、边框、列宽保留），或不套模板新建工作簿 |
| 使用方式 | 命令行引擎 / 本地 Web 操作台 / 纯浏览器零安装 / 打包成 exe、app |
| 效率功能 | 批量转换（多个 A 文件共用一份配置）、配置记忆（按表头指纹自动套用上次配置） |
| 工程特性 | 浏览器模式与 Python 引擎模式**逐值语义对齐**（含数值解析与四舍五入规则） |

## 快速开始

### 方式 A：命令行引擎（保留模板格式，推荐）

```bash
# 依赖：Python 3.9+
pip install -r scripts/requirements.txt      # openpyxl + xlrd
# 或 bash scripts/bootstrap.sh（创建隔离 venv）

python scripts/excel_a2b.py inspect A.xls                 # 1. 看结构：sheet / 表头行 / 列名 / 样例
python scripts/excel_a2b.py inspect B.xlsx --json
python scripts/excel_a2b.py init > mapping.json           # 2. 生成配置模板
python scripts/excel_a2b.py convert --config mapping.json # 3. 执行转换
python scripts/excel_a2b.py presets                       # 查看 15 条内置变换预设
```

也可以完全不写配置文件：

```bash
python scripts/excel_a2b.py convert --source A.xls --template B.xlsx --out out.xlsx \
  --source-header-row 3 --target-header-row 2 \
  --map "下单明细.数量=数量" --map "下单明细.序号=@index" \
  --const "下单明细.产品分类=抓夹"
```

### 方式 B：Web 操作台（人工最省事）

```bash
python web/serve.py            # 自动打开 http://127.0.0.1:8765/
```

页面会自动检测到本地服务并切换到**本地引擎模式**，由 Python 执行转换，完整保留模板格式。
也可以**直接双击 `web/index.html`**（`file://` 方式，零安装），此时由浏览器内的 SheetJS 完成转换。

操作：拖入 A 表和 B 模板 → 自动识别表头行 / 按同名自动匹配 → 可视化编辑映射（变换列可点「预设 ▾」选常用公式）→ 预览（被过滤的行灰显）→ 转换下载。支持导入 / 导出 `mapping.json`，与命令行完全互通。

> 两种模式的区别：纯浏览器模式能保留表头内容、列宽与合并单元格，但**不保留单元格字体 / 边框等样式**；对格式有严格要求请用本地引擎或命令行。

### 方式 C：打包成可执行程序（分享给不懂命令行的人）

```bash
bash web/build_mac.sh        # macOS → dist/excel-a2b
bash scripts/build_windows.sh  # Windows (Git Bash) → dist/excel-a2b.exe
```

打包产物自带 Python 引擎与操作页，双击即用、无需安装 Python。PyInstaller 仅用于打包，不包含在运行时依赖里。

## 目录结构

```
.
├── scripts/
│   ├── excel_a2b.py          # 转换引擎 + CLI（inspect / convert / init / presets）
│   ├── requirements.txt      # openpyxl>=3.1.0, xlrd>=2.0.1
│   ├── bootstrap.sh          # 一键创建隔离 venv
│   └── build_windows.sh      # Windows 打包脚本
├── web/
│   ├── index.html            # 操作台页面
│   ├── serve.py              # 本地引擎服务（静态站点 + /api/convert）
│   ├── build_mac.sh          # macOS 打包脚本
│   ├── assets/               # app.js（全部前端逻辑）、style.css
│   └── vendor/               # SheetJS 0.20.3（离线可用，Apache-2.0）
├── examples/mapping.example.json
├── references/config-schema.md   # 配置字段完整说明
└── docs/WORKBUDDY_SKILL.md        # 作为 WorkBuddy Skill 安装时的说明
```

## 配置示例

```jsonc
{
  "source": { "path": "A.xls", "sheet": "1", "header_row": 3 },
  "target": { "template": "B.xlsx", "path": "输出.xlsx", "sheet": "Sheet1", "header_row": 2 },
  "mapping": [
    { "target": "下单明细.序号", "source": "@index" },
    { "target": "下单明细.数据标题", "source": "{款式}-{货号}", "type": "template" },
    { "target": "下单明细.选择产品货号", "source": "货号" },
    { "target": "下单明细.数量", "source": "数量", "transforms": ["number"] },
    { "target": "下单明细.销售金额", "type": "expr", "expr": "num(数量) * num(单价)" },
    { "target": "下单明细.成本金额", "source": "总价", "transforms": "@金额两位" },
    { "target": "下单明细.产品分类", "type": "const", "value": "抓夹" }
  ],
  "options": {
    "row_filter": {
      "combine": "and",
      "rules": [{ "column": "货号", "match": "^SH\\d+C-\\d+$", "mode": "include" }]
    },
    "split": null,
    "aggregate": null
  }
}
```

行加工顺序：**过滤 → 拆行 → 聚合 → 字段映射**。

### 常用能力速查

| 需求 | 配置 |
|---|---|
| 跳过末尾 `TOTAL` / 备注行 | `options.row_filter`，多条件 `combine: "and" \| "or"`，`mode: include / exclude` |
| 一行按分隔符拆成多行 | `options.split: {"column":"颜色","delimiter":"/"}` |
| 多行汇总成一行 | `options.aggregate: {"group_by":["货号"],"sum":["数量"],"count_into":"行数"}` |
| 表达式计算 | `{"type":"expr","expr":"num(数量)*num(单价)"}`，中文列名可直接当变量，任意列用 `row["列名"]` |
| 常用变换预设 | 配置里写 `"transforms": "@金额两位"`，或操作台点「预设 ▾」下拉 |

内置 15 条预设：去空格、转数字、金额两位、取整、转大写 / 小写 / 文本、标准日期、日期+30天、首行、补零4位、去 SH 前缀、是/否、加前缀 P-、清空变换。运行 `presets --json` 可得到机器可读清单。

完整字段、16 种变换、自定义预设写法见 [`references/config-schema.md`](references/config-schema.md)。

## 常见问题

- **表头不在第一行**：用 `inspect` 确认 `header_row`（1 基）；操作台也有「自动识别表头行」按钮。双层表头（编码行 + 标签行）通常取标签行。
- **源表是老 `.xls`**：由 `xlrd` 读取；输出统一为 `.xlsx`。
- **为什么预览和结果差 0.01**：已修复。引擎与浏览器统一采用「先 `number` 后 `round`」、四舍五入（正数半值进位、负数按 JS `Math.round` 语义），双端逐值对齐。
- **模板原有数据行**：默认清空（`clear_existing_data: true`），表头行保留。
- **只想给同事用、不想装环境**：用打包脚本生成 exe / app，或直接把 `web/index.html` 发过去双击打开。

## 已知限制

- `round` 的参数 `n` 应为**整数位数**；非整数 `n`（如 `{"round":2.5}`）不在支持范围，两端可能相差末位 1 ULP。内置预设的 `n` 全为整数。
- 负数半值按 JS `Math.round` 语义（并列取更接近 +∞，如 `-0.125 → -0.12`），不是严格财务 half-up。
- `round` 直接作用于带千分位逗号的文本会截断（`1,234.567 → 1`）；预设 `@金额两位` 因先执行 `number` 已规避。
- 纯浏览器模式（SheetJS）不保留字体 / 边框等单元格样式；本地引擎模式保留。
- 自定义变换预设别名只在 Python 引擎侧展开；纯浏览器模式请写展开后的变换链。

## 第三方依赖与许可

见 [THIRD-PARTY.md](THIRD-PARTY.md)。简要：`openpyxl`（MIT）、`xlrd`（BSD-3-Clause）、vendored `SheetJS 0.20.3`（Apache-2.0）、`PyInstaller`（仅打包时使用）。
