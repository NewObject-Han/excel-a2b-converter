#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
excel_a2b.py — 通用 Excel A→B 转换引擎（配置驱动）

设计目标
--------
把「A 表（源表）」按用户自定义的字段对应关系，转换写入「B 表（目标表）」。
- 不做任何业务假设：A→B 的字段对应关系完全由配置文件 / 命令行给出。
- 既能被 AI 工具调用（inspect 出结构 → 生成 mapping.json → convert），
  也能人工直接使用（写一份 config.json，或纯命令行 -m 参数）。

子命令
------
  inspect  打印某个 Excel/CSV 的结构（sheet、表头候选行、列名、样例行），输出可读文本或 JSON。
  convert  按配置执行 A→B 转换。
  init     生成一份带注释的 mapping.json 模板，供人工填写。
  presets  打印内置变换预设（常用公式）清单，支持 --json。

用法示例
--------
  # 1) 看结构（AI / 人工都用它来确认表头行、列名）
  python excel_a2b.py inspect A.xls
  python excel_a2b.py inspect B.xlsx --json

  # 2) 配置驱动转换
  python excel_a2b.py convert --config mapping.json

  # 3) 纯命令行转换（人工快速使用）
  python excel_a2b.py convert --source A.xls --template B.xlsx --out out.xlsx \
      --source-header-row 3 --target-header-row 2 \
      --map "下单明细.数量=数量" --map "下单明细.序号=@index" \
      --const "下单明细.产品分类=抓夹"

  # 4) 生成配置模板
  python excel_a2b.py init > mapping.json

  # 5) 查看内置变换预设（常用公式）
  python excel_a2b.py presets
  python excel_a2b.py presets --json
"""
from __future__ import annotations

import argparse
import csv
import datetime as _dt
import json
import math
import os
import re
import sys
from decimal import Decimal, ROUND_HALF_DOWN, ROUND_HALF_UP
from typing import Any

# --------------------------------------------------------------------------- #
# 可选依赖
# --------------------------------------------------------------------------- #
try:
    import openpyxl
    from openpyxl.styles import PatternFill
    from openpyxl.utils import get_column_letter, column_index_from_string
except Exception:  # pragma: no cover
    openpyxl = None

try:
    import xlrd
except Exception:  # pragma: no cover
    xlrd = None


# --------------------------------------------------------------------------- #
# 通用工具
# --------------------------------------------------------------------------- #
def _is_blank(v: Any) -> bool:
    if v is None:
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False


def _to_text(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (_dt.datetime, _dt.date)):
        return v.strftime("%Y-%m-%d")
    return str(v)


def _to_num(v: Any) -> float:
    """转数值（去千分位逗号），失败返回 0.0。供聚合与 expr 使用。"""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    try:
        return float(_to_text(v).replace(",", ""))
    except Exception:
        return 0.0


def _num_clean(f: float) -> Any:
    """整数值转 int，避免写出 3.0。"""
    return int(f) if float(f).is_integer() else f


# expr 表达式的安全求值环境：无内建函数，仅提供白名单函数 + 行列变量
_EXPR_FUNCS = {
    "float": float, "int": int, "str": str, "round": round,
    "len": len, "min": min, "max": max, "abs": abs, "sum": sum, "bool": bool,
    "num": _to_num, "text": _to_text,
}


def eval_expr(expr: str, row: dict[str, Any]) -> Any:
    """计算 expr 表达式。

    - 源表列名若为合法标识符（中文也算），可直接当变量用：``num(数量) * num(单价)``
    - 任意列都可通过 ``row["列名"]`` 访问
    - 无 __builtins__，不能 import / 打开文件，仅白名单函数
    """
    env = dict(_EXPR_FUNCS)
    for k, v in row.items():
        if isinstance(k, str) and k.isidentifier() and k not in env:
            env[k] = v
    env["row"] = row
    return eval(expr, {"__builtins__": {}}, env)


def col_letter_to_index(letter: str) -> int:
    """Excel 列字母 -> 0 基下标。'A'->0, 'AA'->26"""
    return column_index_from_string(letter.upper()) - 1


def index_to_col_letter(idx0: int) -> str:
    return get_column_letter(idx0 + 1)


# --------------------------------------------------------------------------- #
# 读取源表：统一返回 每张 sheet 的 {name, header_row(1基), headers[], rows[[]]}
# --------------------------------------------------------------------------- #
class Table:
    def __init__(self, name: str, matrix: list[list[Any]], header_row: int = 1,
                 data_start_row: int | None = None, end_row: int | None = None):
        self.name = name
        self.matrix = matrix                       # 0 基二维
        self.header_row = header_row               # 1 基
        self.data_start_row = data_start_row or (header_row + 1)
        self.end_row = end_row                     # 1 基，含
        self.headers = self._build_headers()

    def _build_headers(self) -> list[str]:
        r = self.header_row - 1
        if r < 0 or r >= len(self.matrix):
            return []
        raw = self.matrix[r]
        out, seen = [], {}
        for i, v in enumerate(raw):
            name = _to_text(v).strip()
            if name == "":
                name = f"__COL{i + 1}__"          # 无名列占位
            if name in seen:
                seen[name] += 1
                name = f"{name}#{seen[name]}"
            else:
                seen[name] = 0
            out.append(name)
        return out

    @property
    def rows(self) -> list[list[Any]]:
        """数据行（表头行之后到 end_row，去完全空行）。"""
        start = self.data_start_row - 1
        stop = self.end_row if self.end_row else len(self.matrix)
        data = self.matrix[start:stop] if start < len(self.matrix) else []
        return [row for row in data if not all(_is_blank(c) for c in row)]

    def row_dicts(self) -> list[dict[str, Any]]:
        res = []
        for row in self.rows:
            d = {}
            for i, h in enumerate(self.headers):
                d[h] = row[i] if i < len(row) else None
            res.append(d)
        return res

    def col_index(self, spec: Any) -> int | None:
        """列定位：列名 / 1 基序号(int) / Excel 字母(str, 全大写字母) / 'A'.."""
        if spec is None:
            return None
        if isinstance(spec, int) and not isinstance(spec, bool):
            return spec - 1 if spec >= 1 else spec
        s = str(spec).strip()
        if s == "":
            return None
        if s.isdigit():
            n = int(s)
            return n - 1 if n >= 1 else n
        if re.fullmatch(r"[A-Za-z]{1,3}", s):
            return col_letter_to_index(s)
        # 列名精确匹配
        if s in self.headers:
            return self.headers.index(s)
        # 列名包含匹配（唯一时）
        hits = [i for i, h in enumerate(self.headers) if s in h]
        if len(hits) == 1:
            return hits[0]
        return None

    def column(self, spec: Any) -> list[Any]:
        idx = self.col_index(spec)
        if idx is None:
            return []
        return [row[idx] if idx < len(row) else None for row in self.rows]


def _sheet_names(path: str) -> list[str]:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return ["(csv)"]
    if ext == ".xls":
        _require(xlrd, "xlrd", "读取 .xls 需要 xlrd")
        return xlrd.open_workbook(path, on_demand=True).sheet_names()
    _require(openpyxl, "openpyxl", "读取 .xlsx 需要 openpyxl")
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    names = list(wb.sheetnames)
    wb.close()
    return names


def read_table(path: str, sheet: Any = None, header_row: int = 1,
               data_start_row: int | None = None,
               end_row: int | None = None) -> Table:
    """读取任意 sheet 为 Table。header_row 为 1 基。"""
    if not os.path.isfile(path):
        raise FileNotFoundError(f"源文件不存在: {path}")
    ext = os.path.splitext(path)[1].lower()

    if ext == ".csv":
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            matrix = [list(r) for r in csv.reader(f)]
        return Table("(csv)", matrix, header_row, data_start_row, end_row)

    if ext == ".xls":
        _require(xlrd, "xlrd", "读取 .xls 需要 xlrd")
        wb = xlrd.open_workbook(path)
        sh = wb.sheet_by_name(sheet) if isinstance(sheet, str) else \
            wb.sheet_by_index(sheet if isinstance(sheet, int) else 0)
        matrix = []
        for r in range(sh.nrows):
            row = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                val = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        val = xlrd.xldate.xldate_as_datetime(val, wb.datemode)
                    except Exception:
                        pass
                row.append(val)
            matrix.append(row)
        return Table(sh.name, matrix, header_row, data_start_row, end_row)

    # xlsx / xlsm
    _require(openpyxl, "openpyxl", "读取 Excel 需要 openpyxl")
    wb = openpyxl.load_workbook(path, data_only=True)
    sh = wb[sheet] if isinstance(sheet, str) else wb.worksheets[
        sheet if isinstance(sheet, int) else 0]
    matrix = [list(r) for r in sh.iter_rows(values_only=True)]
    wb.close()
    return Table(sh.title, matrix, header_row, data_start_row, end_row)


def _require(mod, name, msg):
    if mod is None:
        raise RuntimeError(f"{msg}（pip install {name}）")


# --------------------------------------------------------------------------- #
# 数据变换（transforms）
# --------------------------------------------------------------------------- #
_TPL_VAR = re.compile(r"\{([^{}]+)\}")

# 内置变换预设：与 Web 操作台「变换」列下拉一一对应。
# 键为 @名称；值为等价的 transforms 列表（元素为字符串或单键字典）。
# 在 mapping 规则里写 "transforms": "@名称" 即可套用；也可用
# options.transform_presets / 顶层 transform_presets 覆盖或追加。
TRANSFORM_PRESETS: dict[str, list[Any]] = {
    "@去空格": ["trim"],
    "@转数字": ["trim", "number"],
    "@金额两位": ["number", {"round": 2}],
    "@取整": ["number", "int"],
    "@转大写": ["trim", "upper"],
    "@转小写": ["trim", "lower"],
    "@转文本": ["text"],
    "@标准日期": [{"date": "%Y-%m-%d"}],
    "@日期+30天": [{"date_offset": 30}],
    "@首行": ["first_line"],
    "@补零4位": [{"pad": ["0", 4]}],
    "@去SH前缀": [{"regex_replace": ["^SH", ""]}],
    "@是/否": ["bool"],
    "@加前缀P-": [{"prefix": "P-"}],
    "@清空变换": [],
}


def _preset_item_text(spec: Any) -> str:
    """把单个变换 spec 渲染成操作台「变换」输入框的等价文本。"""
    if isinstance(spec, str):
        return spec
    compact = dict(ensure_ascii=False, separators=(",", ":"))
    if isinstance(spec, dict) and len(spec) == 1:
        name, arg = next(iter(spec.items()))
        return f"{name}:{json.dumps(arg, **compact)}"
    return json.dumps(spec, **compact)


def _preset_chain_text(chain: list[Any]) -> str:
    """把一条预设链渲染成逗号分隔文本（与 Web 预设 value 完全一致）。"""
    return ",".join(_preset_item_text(t) for t in chain)


def merge_transform_presets(cfg: dict) -> dict[str, list[Any]]:
    """合并内置预设与自定义预设，返回「别名 -> 变换链列表」。

    自定义预设可写在配置顶层 ``transform_presets`` 或
    ``options.transform_presets``；后者同名优先。键可省略 ``@`` 前缀。
    值可以是变换链数组、单个变换名字符串或单个变换字典。
    """
    merged = {k: list(v) for k, v in TRANSFORM_PRESETS.items()}
    custom: dict[str, Any] = {}
    for source in (cfg.get("transform_presets"),
                   (cfg.get("options") or {}).get("transform_presets")):
        if isinstance(source, dict):
            custom.update(source)
    for key, chain in custom.items():
        name = str(key)
        if not name.startswith("@"):
            name = "@" + name
        if chain is None:
            merged[name] = []
        elif isinstance(chain, (str, dict)):
            merged[name] = [chain]
        elif isinstance(chain, (list, tuple)):
            merged[name] = list(chain)
        else:
            merged[name] = []
    return merged


def expand_transforms(rule: dict,
                      presets: dict[str, list[Any]] | None = None) -> list[Any]:
    """把一条规则的 transforms 展开成可逐项执行的列表。

    - 字符串：``"@名称"`` 查预设表展开；查不到则按**单个变换名**处理
      （修复旧实现里字符串被逐字符迭代的问题）。
    - 列表：逐项处理，元素本身也可以是 ``"@名称"`` 预设别名。
    - 单个字典：包成列表。
    """
    table = presets if presets is not None else TRANSFORM_PRESETS
    raw = rule.get("transforms", []) or []
    if isinstance(raw, str):
        key = raw.strip()
        if not key:
            return []
        return list(table[key]) if key in table else [key]
    if isinstance(raw, dict):
        return [raw]
    if isinstance(raw, (list, tuple)):
        out: list[Any] = []
        for spec in raw:
            if isinstance(spec, str):
                key = spec.strip()
                if key in table:
                    out.extend(list(table[key]))
                elif key:
                    out.append(key)
            else:
                out.append(spec)
        return out
    return [raw]


# 数值前缀：对齐浏览器 parseFloat / parseInt 的「宽容数值前缀」语义。
# 非法后缀（如 "12abc"）被忽略；无匹配（如 "" / "abc"）则视为解析失败。
_NUM_PREFIX = re.compile(r"^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?")
_INT_PREFIX = re.compile(r"^[+-]?\d+")
_INF_TEXT = re.compile(r"^([+-]?)Infinity$")


def _lenient_float(v: Any, *, strip_commas: bool = True,
                   integer_only: bool = False) -> float | None:
    """按 ``parseFloat`` / ``parseInt`` 语义提取数值前缀，失败返回 ``None``。

    - ``strip_commas=True``：对齐 JS ``parseFloat(toText(v).replace(/,/g,''))``
      （``int`` / ``number`` 分支）；``False`` 对齐 ``parseFloat(toText(v))``
      （``round`` 分支在 JS 侧未去千分位逗号，此处如实复刻以保证双端一致）。
    - ``integer_only=True``：对齐 JS ``parseInt(s, 10)``，只认整数前缀，
      不认小数点 / 指数 / 前导点（``"1e3"`` -> 1、``".5"`` -> 解析失败）。
    - 空串 / 纯文本 / ``"abc"`` / 非有限值一律返回 ``None``，由调用点返回原值。
    """
    s = _to_text(v).strip()
    if strip_commas:
        s = s.replace(",", "")
    if not integer_only:
        # parseFloat("Infinity" / "1e999") 返回 Infinity（不是 NaN），如实复刻
        mi = _INF_TEXT.match(s)
        if mi:
            return float("-inf") if mi.group(1) == "-" else float("inf")
    m = (_INT_PREFIX if integer_only else _NUM_PREFIX).match(s)
    if not m:
        return None
    try:
        return float(m.group(0))       # "1e999" 会溢出为 inf，同样与 JS 一致
    except (ValueError, OverflowError):
        return None


def _round_like_js(f: float, n: float) -> Any:
    """复刻浏览器 ``Math.round(f * 10**n) / 10**n`` 的结果。

    先对 ``f * 10**n``（IEEE 754 double，与 JS 计算出的值完全相同）取整，
    因此浮点表示带来的边界行为也一致（如 ``1.005`` -> 1.0）。

    正数半值用 ``Decimal`` + ``ROUND_HALF_UP``（财务四舍五入：
    ``0.125`` -> 0.13、``2.675`` -> 2.68）；负数半值按 JS ``Math.round``
    「并列时取更接近 +∞」的语义改用 ``ROUND_HALF_DOWN``
    （``-0.125`` -> -0.12），从而保证双端同输入同输出。
    """
    if not math.isfinite(f):
        return f                       # JS: Math.round(Infinity) 仍是 Infinity
    m = 10.0 ** n
    scaled = f * m
    mode = ROUND_HALF_UP if scaled >= 0 else ROUND_HALF_DOWN
    q = Decimal(str(scaled)).quantize(Decimal(1), rounding=mode)
    result = float(q) / m
    return int(result) if result.is_integer() else result


def apply_transform(value: Any, spec: Any) -> Any:
    """单项变换。spec 为字符串或单键字典。"""
    if isinstance(spec, dict):
        name, arg = next(iter(spec.items()))
    else:
        name, arg = str(spec), None
    name = name.strip().lower()

    if name in ("trim", "strip"):
        return _to_text(value).strip()
    if name == "upper":
        return _to_text(value).upper()
    if name == "lower":
        return _to_text(value).lower()
    if name == "text":
        return _to_text(value)
    if name == "int":
        # 对齐 JS parseInt(s, 10)：只认整数前缀，指数 / 前导点不认
        f = _lenient_float(value, integer_only=True)
        return int(f) if f is not None else value
    if name == "number":
        # 对齐 JS parseFloat(s)：宽容数值前缀，整数值写成 int 避免 12.0
        f = _lenient_float(value)
        if f is None:
            return value
        return int(f) if f.is_integer() else f
    if name == "round":
        # 对齐 JS Math.round(f * 10^n) / 10^n：Decimal half-up（非银行家舍入）
        # JS 的 round 分支未去千分位逗号，这里 strip_commas=False 如实复刻。
        f = _lenient_float(value, strip_commas=False)
        if f is None:
            return value
        try:
            raw_n = arg[0] if isinstance(arg, (list, tuple)) and arg else (arg or 0)
            # JS 用 +(arg||0) 参与 Math.pow，非整数位数（2.5）也必须如实保留
            return _round_like_js(f, float(raw_n))
        except Exception:
            return value
    if name == "date":
        return _fmt_date(value, str(arg or "%Y-%m-%d"))
    if name == "date_offset":
        d = _parse_date(value)
        if d is None:
            return value
        return (d + _dt.timedelta(days=int(arg or 0))).strftime("%Y-%m-%d")
    if name == "prefix":
        return f"{arg}{_to_text(value)}"
    if name == "suffix":
        return f"{_to_text(value)}{arg}"
    if name == "replace":
        pat, repl = (arg or ["", ""])[0], (arg or ["", ""])[1]
        return re.sub(pat, repl, _to_text(value))
    if name == "regex_replace":
        pat, repl = (arg or ["", ""])[0], (arg or ["", ""])[1]
        return re.sub(pat, repl, _to_text(value))
    if name == "split":
        sep, idx = (arg or ["\n", 0])[0], int((arg or ["\n", 0])[1])
        parts = _to_text(value).split(sep)
        return parts[idx].strip() if -len(parts) <= idx < len(parts) else ""
    if name == "first_line":
        return _to_text(value).splitlines()[0].strip() if _to_text(value) else ""
    if name == "pad":
        fill, width = (arg or ["0", 0])[0], int((arg or ["0", 0])[1])
        return _to_text(value).rjust(width, fill)
    if name == "bool":
        return 1 if _to_text(value).strip() not in ("", "0", "false", "False", "否") else 0
    return value


def _parse_date(v: Any) -> _dt.datetime | None:
    if isinstance(v, _dt.datetime):
        return v
    if isinstance(v, _dt.date):
        return _dt.datetime(v.year, v.month, v.day)
    if isinstance(v, (int, float)):
        try:
            return _dt.datetime(1899, 12, 30) + _dt.timedelta(days=float(v))
        except Exception:
            return None
    s = _to_text(v).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日",
                "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"):
        try:
            return _dt.datetime.strptime(s, fmt)
        except Exception:
            continue
    return None


def _fmt_date(v: Any, fmt: str) -> str:
    d = _parse_date(v)
    return d.strftime(fmt) if d else _to_text(v)


def render_template(tpl: str, row: dict[str, Any], row_index: int) -> str:
    """把 '{列名}' 占位符替换为当前行的值。支持 {#index}。"""
    def repl(m):
        key = m.group(1).strip()
        if key in ("#index", "index", "序号"):
            return str(row_index)
        return _to_text(row.get(key, ""))
    return _TPL_VAR.sub(repl, tpl)


# --------------------------------------------------------------------------- #
# 单元格取值（一条 mapping 规则 -> 一个值）
# --------------------------------------------------------------------------- #
def _raw_cell(row: dict[str, Any], table: Table, src_spec: Any) -> Any:
    """按列名/索引取源值。"""
    idx = table.col_index(src_spec)
    if idx is not None and idx < len(table.headers):
        return row.get(table.headers[idx])
    # 源列名不在表头：直接按名字查（可能来自 __COLx__）
    if isinstance(src_spec, str) and src_spec in row:
        return row[src_spec]
    return None


def resolve_value(rule: dict, row: dict, table: Table, row_index: int,
                  on_missing: str = "empty") -> Any:
    """根据一条 mapping 规则与当前行，计算目标值。"""
    rtype = str(rule.get("type", "")).strip().lower()

    # 1) 常量
    if rtype == "const" or ("value" in rule and "source" not in rule):
        return rule.get("value", "")

    # 2) 行号
    if rtype in ("index", "row_number") or str(rule.get("source", "")).strip().lower() in ("@index", "#index"):
        start = rule.get("start", 0)
        return row_index + int(start)

    # 3) 表达式（列名作变量，如 num(数量) * num(单价)）
    if rtype == "expr":
        expr = str(rule.get("expr") or rule.get("source") or "")
        if not expr.strip():
            return ""
        return eval_expr(expr, row)

    src = rule.get("source")

    # 4) 模板拼接（{列}）
    if rtype == "template" or (isinstance(src, str) and _TPL_VAR.search(src)):
        return render_template(str(src), row, row_index)

    # 5) 普通取列
    val = _raw_cell(row, table, src)
    if val is None and src is not None and _to_text(src) not in table.headers:
        if on_missing == "error":
            raise KeyError(f"源列不存在: {src!r}（可用列：{table.headers}）")
        if on_missing == "keep":
            return str(src)
    return val


def finalize(value: Any, rule: dict,
             presets: dict[str, list[Any]] | None = None) -> Any:
    """应用默认值 + 变换链。

    ``presets`` 为合并后的预设表（见 :func:`merge_transform_presets`）；
    不传时只用内置预设，保持 ``finalize(value, rule)`` 的既有调用兼容。
    """
    if _is_blank(value):
        if "default" in rule:
            value = rule["default"]
    for t in expand_transforms(rule, presets):
        value = apply_transform(value, t)
    return value


# --------------------------------------------------------------------------- #
# 目标写入
# --------------------------------------------------------------------------- #
def _resolve_target_col(ws, header_row: int, spec: Any,
                        create_missing: bool) -> tuple[int | None, str]:
    """在目标 sheet 的表头行定位列。返回 (1 基列号, 列名)。"""
    # 表头行内容
    headers = {}
    if header_row and header_row >= 1:
        for c in range(1, ws.max_column + 2):
            v = _to_text(ws.cell(row=header_row, column=c).value).strip()
            if v:
                headers.setdefault(v, c)
    if isinstance(spec, int) and not isinstance(spec, bool):
        return spec, headers.get(spec, str(spec))
    s = str(spec).strip()
    if s.isdigit():
        return int(s), headers.get(s, s)
    if re.fullmatch(r"[A-Za-z]{1,3}", s):
        return column_index_from_string(s.upper()), s
    if s in headers:
        return headers[s], s
    hits = [c for h, c in headers.items() if s and s in h]
    if len(hits) == 1:
        return hits[0], s
    if create_missing:
        col = (max(headers.values()) if headers else 0) + 1
        if header_row and header_row >= 1:
            ws.cell(row=header_row, column=col).value = s
        return col, s
    raise KeyError(f"目标表表头未找到列: {spec!r}（可用：{list(headers)}）")


def _normalize_row_filter(rf: Any) -> dict:
    """把多种写法统一成 {combine, rules:[...]}。

    - 列表：视为多条规则，combine 默认 and
    - 含 rules 字段的对象：取 combine / rules
    - 单条规则对象（旧写法）：包成一条
    """
    if isinstance(rf, list):
        return {"combine": "and", "rules": rf}
    if isinstance(rf, dict):
        if "rules" in rf:
            return {"combine": rf.get("combine", "and"), "rules": rf.get("rules") or []}
        return {"combine": "and", "rules": [rf]}
    return {"combine": "and", "rules": []}


def convert(cfg: dict) -> dict:
    """执行转换，返回统计信息。"""
    src_cfg = cfg.get("source", {})
    tgt_cfg = cfg.get("target", {})
    opt = cfg.get("options", {})
    mapping = cfg.get("mapping", [])
    if not mapping:
        raise ValueError("配置里 mapping 为空，没有可转换的字段")

    # 合并内置 + 自定义变换预设，供每条规则的 finalize 使用
    presets = merge_transform_presets(cfg)

    on_missing = opt.get("on_missing_source", "empty")
    skip_empty = opt.get("skip_empty_source_rows", True)

    # ---- 读源表 ----
    table = read_table(
        src_cfg["path"],
        sheet=src_cfg.get("sheet"),
        header_row=int(src_cfg.get("header_row") or 1),
        data_start_row=src_cfg.get("data_start_row") or None,
        end_row=src_cfg.get("end_row") or None,
    )
    src_rows = table.row_dicts()
    n_src_total = len(src_rows)

    # ---- 行过滤（按列格式判断，跳过非数据行，如末尾的 TOTAL / 客户要求）----
    # 支持「多个条件」：
    #   单条（旧写法，向后兼容）：
    #     "row_filter": {"column":"货号","match":"^SH\\d+C-\\d+$","mode":"include"}
    #   多条（数组，combine=and/or）：
    #     "row_filter": {"combine":"and","rules":[
    #        {"column":"货号","match":"^SH\\d+C-\\d+$","mode":"include"},
    #        {"column":"数量","match":"^\\d+$","mode":"include"}]}
    #   - mode=include（默认）：保留「该列值匹配正则」的行
    #   - mode=exclude        ：跳过「该列值匹配正则」的行
    #   - combine=and：所有条件都满足才保留；combine=or：任一满足即保留
    row_filter = opt.get("row_filter") or cfg.get("row_filter")
    n_filtered = 0
    if row_filter:
        norm = _normalize_row_filter(row_filter)
        combine = str(norm.get("combine", "and")).strip().lower()
        compiled = []
        for r in norm.get("rules", []) or []:
            rc = r.get("column") or r.get("source")
            rp = r.get("match") or r.get("pattern") or r.get("regex")
            rm = str(r.get("mode", "include")).strip().lower()
            if not (rc and rp):
                continue
            try:
                rx = re.compile(rp)
            except re.error as e:
                raise ValueError(f"row_filter 正则非法：{rp}（{e}）")
            compiled.append((rc, rx, rm))
        if compiled:
            kept = []
            for row in src_rows:
                res = []
                for (rc, rx, rm) in compiled:
                    val = _to_text(row.get(rc, ""))
                    is_match = bool(rx.search(val))
                    res.append(is_match if rm != "exclude" else not is_match)
                keep = all(res) if combine == "and" else any(res)
                if keep:
                    kept.append(row)
                else:
                    n_filtered += 1
            src_rows = kept

    # ---- 拆行（一行 A → 多行 B）：把某列按分隔符拆开 ----
    # 配置（options.split 或 source.split）：
    #   {"column":"颜色","delimiter":"/","regex":false,"trim":true,"skip_empty":true}
    split_cfg = opt.get("split") or src_cfg.get("split")
    n_before_split = len(src_rows)
    if split_cfg:
        sp_col = split_cfg.get("column")
        sp_delim = str(split_cfg.get("delimiter", "/"))
        sp_regex = bool(split_cfg.get("regex"))
        sp_trim = split_cfg.get("trim", True)
        sp_skip_empty = split_cfg.get("skip_empty", True)
        if sp_col:
            new_rows = []
            for row in src_rows:
                val = _to_text(row.get(sp_col, ""))
                parts = re.split(sp_delim, val) if sp_regex else val.split(sp_delim)
                for p in parts:
                    p2 = p.strip() if sp_trim else p
                    if sp_skip_empty and p2 == "":
                        continue
                    r2 = dict(row)
                    r2[sp_col] = p2
                    new_rows.append(r2)
            src_rows = new_rows

    # ---- 聚合（多行 A → 一行 B）：按分组列汇总，指定列求和 ----
    # 配置（options.aggregate 或 source.aggregate）：
    #   {"group_by":["货号"],"sum":["数量","总价"],"count_into":"行数"}
    # 未列入 sum 的列取组内第一行的值；count_into 可选，把组内行数写入该列名。
    agg_cfg = opt.get("aggregate") or src_cfg.get("aggregate")
    n_before_agg = len(src_rows)
    if agg_cfg:
        group_by = agg_cfg.get("group_by") or []
        if isinstance(group_by, str):
            group_by = [group_by]
        sum_cols = agg_cfg.get("sum") or []
        count_into = agg_cfg.get("count_into")
        if group_by:
            groups: dict[tuple, dict] = {}
            order: list[tuple] = []
            for row in src_rows:
                key = tuple(_to_text(row.get(g, "")) for g in group_by)
                g = groups.get(key)
                if g is None:
                    g = dict(row)
                    for sc in sum_cols:
                        g[sc] = _to_num(row.get(sc))
                    g["__group_count"] = 1
                    groups[key] = g
                    order.append(key)
                else:
                    for sc in sum_cols:
                        g[sc] = _to_num(g.get(sc)) + _to_num(row.get(sc))
                    g["__group_count"] += 1
            src_rows = []
            for key in order:
                g = groups[key]
                for sc in sum_cols:
                    g[sc] = _num_clean(g[sc])
                if count_into:
                    g[count_into] = g["__group_count"]
                g.pop("__group_count", None)
                src_rows.append(g)

    # ---- 逐行计算目标值 ----
    records = []
    for i, row in enumerate(src_rows, start=1):
        if skip_empty and all(_is_blank(v) for v in row.values()):
            continue
        rec = {}
        for rule in mapping:
            try:
                val = resolve_value(rule, row, table, i, on_missing)
                rec[rule["target"]] = finalize(val, rule, presets)
            except Exception as e:
                raise ValueError(f"目标列 {rule.get('target')!r} 第 {i} 行取值失败：{e}") from e
        records.append(rec)

    out_path = tgt_cfg.get("path")
    if not out_path:
        raise ValueError("target.path 必填（输出文件路径）")

    # ---- 写目标 ----
    stats = {"source_rows": len(src_rows),
             "source_rows_before_filter": n_src_total,
             "skipped_by_filter": n_filtered,
             "written_rows": len(records),
             "output": out_path, "target_columns": len(mapping)}
    if split_cfg:
        stats["rows_before_split"] = n_before_split
    if agg_cfg:
        stats["rows_before_aggregate"] = n_before_agg

    template = tgt_cfg.get("template")
    if template:
        _require(openpyxl, "openpyxl", "写入 xlsx 需要 openpyxl")
        wb = openpyxl.load_workbook(template)
        ws = wb[tgt_cfg["sheet"]] if tgt_cfg.get("sheet") else wb.worksheets[0]
        header_row = int(tgt_cfg.get("header_row") or 1)
        data_start = int(tgt_cfg.get("data_start_row") or (header_row + 1))
        clear = opt.get("clear_existing_data", True)

        # 预解析目标列（一次性），避免逐行重复扫描表头
        resolved = []
        meta_by_target = {}
        for rule in mapping:
            col, name = _resolve_target_col(ws, header_row, rule["target"],
                                            opt.get("create_missing_columns", False))
            resolved.append((rule, col))
            meta_by_target[rule["target"]] = {"column": col, "column_letter": index_to_col_letter(col - 1)}

        # 记录首个数据行的样式，供后续行复制
        style_probe = {}
        for rule, col in resolved:
            style_probe[col] = ws.cell(row=data_start, column=col)

        if clear:
            for r in range(data_start, ws.max_row + 1):
                for c in range(1, ws.max_column + 1):
                    ws.cell(row=r, column=c).value = None

        for ri, rec in enumerate(records):
            r = data_start + ri
            for rule, col in resolved:
                cell = ws.cell(row=r, column=col)
                cell.value = rec[rule["target"]]
                src_cell = style_probe.get(col)
                if src_cell is not None and src_cell.has_style:
                    cell._style = src_cell._style
        wb.save(out_path)
        stats["target_columns_detail"] = meta_by_target
    else:
        # 不套模板：新建工作簿，表头 = mapping 的 target
        _require(openpyxl, "openpyxl", "写入 xlsx 需要 openpyxl")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = tgt_cfg.get("sheet") or "Sheet1"
        for c, rule in enumerate(mapping, start=1):
            ws.cell(row=1, column=c).value = rule["target"]
        for ri, rec in enumerate(records):
            for c, rule in enumerate(mapping, start=1):
                ws.cell(row=2 + ri, column=c).value = rec[rule["target"]]
        wb.save(out_path)

    return stats


# --------------------------------------------------------------------------- #
# inspect
# --------------------------------------------------------------------------- #
def inspect_file(path: str, sheet: Any = None, preview: int = 5,
                 scan_header_rows: int = 8, header_row: int | None = None) -> dict:
    result: dict[str, Any] = {"file": os.path.abspath(path), "sheets": []}
    names = _sheet_names(path)

    targets = [sheet] if sheet is not None else names
    for name in targets:
        # 直接取原始 matrix（自己扫表头行）
        t = read_table(path, sheet=name, header_row=1)
        m = t.matrix
        cand = []
        for r in range(min(scan_header_rows, len(m))):
            filled = sum(1 for v in m[r] if not _is_blank(v))
            cand.append({"row": r + 1, "non_empty": filled,
                         "sample": [_to_text(v) for v in m[r][:12]]})

        # 表头推断：取「非空单元格最多的候选行中的最后一行」。
        # 这样可正确处理「机器编码行 + 人工标签行」双层表头（取下面的标签行）。
        max_ne = max((c["non_empty"] for c in cand), default=0)
        top = [c for c in cand if c["non_empty"] == max_ne]
        best = header_row or (top[-1]["row"] if top else 1)

        hint = None
        if len(top) >= 2 and not header_row:
            hint = (f"检测到第 {top[0]['row']} 行与第 {top[-1]['row']} 行均为满行，"
                    f"可能是「编码行+标签行」双层表头，已建议取标签行。")

        info = {
            "name": name,
            "rows": len(m),
            "cols": max((len(r) for r in m), default=0),
            "header_candidates": cand,
            "suggested_header_row": best,
            "hint": hint,
            "columns": [],
            "preview": [],
        }
        t2 = Table(name, m, best)
        info["columns"] = t2.headers
        for row in t2.rows[:preview]:
            info["preview"].append([_to_text(v) for v in row[:len(t2.headers)]])
        result["sheets"].append(info)
    return result


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def _cmd_inspect(args):
    res = inspect_file(args.file, sheet=args.sheet, preview=args.preview,
                       header_row=args.header_row)
    if args.json:
        print(json.dumps(res, ensure_ascii=False, indent=2))
        return
    for sh in res["sheets"]:
        print("=" * 70)
        print(f"文件: {res['file']}")
        print(f"Sheet: {sh['name']}   行数={sh['rows']}  列数={sh['cols']}")
        print(f"建议表头行: 第 {sh['suggested_header_row']} 行")
        if sh.get("hint"):
            print(f"提示: {sh['hint']}")
        print("-" * 70)
        for c in sh["header_candidates"]:
            mark = " <== 表头" if c["row"] == sh["suggested_header_row"] else ""
            print(f"  第{c['row']}行 (非空{c['non_empty']:>2}): {c['sample']}{mark}")
        print("-" * 70)
        print(f"列名({len(sh['columns'])}): {sh['columns']}")
        print("样例行:")
        for row in sh["preview"]:
            print("   ", row)
    print("=" * 70)


def _cmd_init(args):
    tpl = {
        "source": {
            "path": "A表路径.xls",
            "sheet": None,
            "header_row": 1,
            "data_start_row": None,
        },
        "target": {
            "template": "B表模板路径.xlsx",
            "path": "输出.xlsx",
            "sheet": None,
            "header_row": 1,
            "data_start_row": None,
        },
        "mapping": [
            {"target": "目标列名", "source": "源列名"},
            {"target": "目标列名2", "source": "{列A}-{列B}", "type": "template"},
            {"target": "目标列名3", "source": "@index"},
            {"target": "目标固定列", "type": "const", "value": "固定值"},
            {"target": "目标数值列", "source": "数量", "transforms": ["number"]},
            {"target": "目标日期列", "source": "下单日期", "transforms": "@标准日期"},
            {"target": "目标表达式列", "type": "expr", "expr": "num(数量) * num(单价)"},
        ],
        "options": {
            "skip_empty_source_rows": True,
            "clear_existing_data": True,
            "on_missing_source": "empty",
            "create_missing_columns": False,
            "row_filter": {"column": "货号", "match": "^SH\\d+C-\\d+$", "mode": "include"},
            "split": None,
            "aggregate": None,
        },
    }
    print(json.dumps(tpl, ensure_ascii=False, indent=2))


def _cmd_convert(args):
    if args.config:
        with open(args.config, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    else:
        if not args.source or not args.out:
            raise SystemExit("命令行模式必须提供 --source 与 --out（或改用 --config）")
        mapping = []
        for pair in (args.map or []):
            if "=" not in pair:
                raise SystemExit(f"--map 格式应为 目标列=源列，收到: {pair}")
            t, s = pair.split("=", 1)
            entry = {"target": t.strip()}
            if s.strip().lower() in ("@index", "#index"):
                entry.update({"source": "@index"})
            else:
                entry.update({"source": s.strip()})
            mapping.append(entry)
        for pair in (args.const or []):
            if "=" not in pair:
                raise SystemExit(f"--const 格式应为 目标列=值，收到: {pair}")
            t, v = pair.split("=", 1)
            mapping.append({"target": t.strip(), "type": "const", "value": v})
        if not mapping:
            raise SystemExit("未提供任何映射，请用 --map/--const 或 --config")
        cfg = {
            "source": {
                "path": args.source,
                "sheet": args.source_sheet,
                "header_row": args.source_header_row,
                "data_start_row": args.source_data_start_row,
            },
            "target": {
                "template": args.template,
                "path": args.out,
                "sheet": args.target_sheet,
                "header_row": args.target_header_row,
                "data_start_row": args.target_data_start_row,
            },
            "mapping": mapping,
            "options": {
                "skip_empty_source_rows": not args.keep_empty_rows,
                "clear_existing_data": not args.keep_template_data,
                "created_by": "cli",
            },
        }

    # 覆盖选项
    opts = cfg.setdefault("options", {})
    if getattr(args, "keep_empty_rows", False):
        opts["skip_empty_source_rows"] = False
    if getattr(args, "keep_template_data", False):
        opts["clear_existing_data"] = False
    if getattr(args, "create_missing_columns", False):
        opts["create_missing_columns"] = True
    if getattr(args, "on_missing", None):
        opts["on_missing_source"] = args.on_missing
    if getattr(args, "out", None) and not args.config:
        cfg["target"]["path"] = args.out

    stats = convert(cfg)
    print(json.dumps(stats, ensure_ascii=False, indent=2))


def _cmd_presets(args):
    """打印内置变换预设清单：人类可读表格，或 --json 机器可读。"""
    if args.json:
        print(json.dumps({
            "presets": TRANSFORM_PRESETS,
            "web_values": {k: _preset_chain_text(v) for k, v in TRANSFORM_PRESETS.items()},
        }, ensure_ascii=False, indent=2))
        return
    print("内置变换预设（常用公式）")
    print("在 mapping 规则里写 \"transforms\": \"@别名\" 即可套用；Web 操作台点「预设 ▾」下拉直接选用。")
    print("=" * 74)
    print(f"{'别名':<14}{'等价变换链（Web 变换输入框文本）'}")
    print("-" * 74)
    for key, chain in TRANSFORM_PRESETS.items():
        text = _preset_chain_text(chain) or "（空，一键清空变换）"
        print(f"{key:<14}{text}")
    print("=" * 74)
    print("自定义预设：options.transform_presets / 顶层 transform_presets；用 --json 获取机器可读版本。")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="excel_a2b",
        description="通用 Excel A→B 转换引擎（字段对应关系完全自定义）",
    )
    sub = p.add_subparsers(dest="cmd", required=True)

    pi = sub.add_parser("inspect", help="查看文件结构（sheet/表头候选/列名/样例）")
    pi.add_argument("file")
    pi.add_argument("--sheet", default=None)
    pi.add_argument("--header-row", type=int, default=None, dest="header_row",
                    help="强制指定表头行（1 基）")
    pi.add_argument("--preview", type=int, default=5)
    pi.add_argument("--json", action="store_true")
    pi.set_defaults(func=_cmd_inspect)

    pc = sub.add_parser("convert", help="执行 A→B 转换")
    pc.add_argument("--config", help="mapping.json 配置文件路径")
    pc.add_argument("--source", help="A 表路径")
    pc.add_argument("--template", help="B 表模板路径（可选，用于继承表头/格式）")
    pc.add_argument("--out", help="输出 xlsx 路径")
    pc.add_argument("--map", action="append", help="目标列=源列，可重复")
    pc.add_argument("--const", action="append", help="目标列=固定值，可重复")
    pc.add_argument("--source-sheet", default=None)
    pc.add_argument("--source-header-row", type=int, default=1, dest="source_header_row")
    pc.add_argument("--source-data-start-row", type=int, default=None, dest="source_data_start_row")
    pc.add_argument("--target-sheet", default=None)
    pc.add_argument("--target-header-row", type=int, default=1, dest="target_header_row")
    pc.add_argument("--target-data-start-row", type=int, default=None, dest="target_data_start_row")
    pc.add_argument("--keep-empty-rows", action="store_true", dest="keep_empty_rows")
    pc.add_argument("--keep-template-data", action="store_true", dest="keep_template_data")
    pc.add_argument("--create-missing-columns", action="store_true", dest="create_missing_columns")
    pc.add_argument("--on-missing", choices=["empty", "error", "keep"], default=None)
    pc.set_defaults(func=_cmd_convert)

    pt = sub.add_parser("init", help="打印一份配置模板")
    pt.set_defaults(func=_cmd_init)

    pp = sub.add_parser("presets", help="列出内置变换预设（常用公式）清单")
    pp.add_argument("--json", action="store_true", help="以 JSON 输出，便于程序读取")
    pp.set_defaults(func=_cmd_presets)
    return p


def main(argv=None):
    args = build_parser().parse_args(argv)
    try:
        args.func(args)
    except Exception as e:
        print(f"[ERROR] {type(e).__name__}: {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
