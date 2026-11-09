#!/usr/bin/env bash
# 初始化 excel-a2b-converter 的隔离运行环境（创建 venv 并安装依赖）。
# 用法:  bash bootstrap.sh
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV="${EXCEL_A2B_VENV:-$HOME/.workbuddy/binaries/python/envs/excel-a2b}"

find_python() {
  local cands=()
  [ -n "${EXCEL_A2B_PYTHON:-}" ] && cands+=("$EXCEL_A2B_PYTHON")
  cands+=("$HOME/.workbuddy/binaries/python/versions/3.13.12/python.exe")
  cands+=("python3" "python")
  for c in "${cands[@]}"; do
    if [ -n "$c" ] && command -v "$c" >/dev/null 2>&1; then echo "$c"; return 0; fi
  done
  return 1
}

PY="$(find_python)" || { echo "未找到可用的 Python 解释器，请设置 EXCEL_A2B_PYTHON" >&2; exit 1; }
echo "使用解释器: $PY"

if [ ! -x "$VENV/Scripts/python.exe" ] && [ ! -x "$VENV/bin/python" ]; then
  echo "创建虚拟环境: $VENV"
  "$PY" -m venv "$VENV"
fi

if [ -x "$VENV/Scripts/python.exe" ]; then VPY="$VENV/Scripts/python.exe"; else VPY="$VENV/bin/python"; fi

"$VPY" -m pip install --quiet --upgrade pip
"$VPY" -m pip install --quiet -r "$HERE/requirements.txt"
echo "环境就绪: $VPY"
"$VPY" -c "import openpyxl, xlrd; print('openpyxl', openpyxl.__version__, '| xlrd', xlrd.__version__)"
