#!/usr/bin/env bash
# 在 macOS 上把操作台打包成可执行程序（双击即用，本地引擎保真模板格式）
# 用法： bash build_mac.sh
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"   # web/
SK="$(cd "$HERE/.." && pwd)"
cd "$HERE"

PY="${PYTHON:-python3}"

# 安装依赖（openpyxl / xlrd 纯 Python，可跨平台）
"$PY" -m pip install --quiet pyinstaller openpyxl xlrd || {
  echo "请先安装 Python 3（brew install python / 官网 pkg），再运行本脚本。" >&2
  exit 1
}

# 注意 macOS 的 --add-data 分隔符是冒号 ":"
"$PY" -m PyInstaller --onefile --name excel-a2b --windowed \
  --add-data "$SK/web:web" --add-data "$SK/scripts:scripts" \
  --paths "$SK/scripts" --hidden-import openpyxl --hidden-import xlrd \
  serve.py

echo "完成：dist/excel-a2b 即为 macOS 可执行程序（双击运行，自动打开浏览器）。"
echo "如需更易用的 .app，可在 dist/ 上用 py2app 或手动封装，或直接分发 dist/excel-a2b 也可双击运行。"
