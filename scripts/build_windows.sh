#!/usr/bin/env bash
# 在 Windows（Git Bash / MSYS）上把操作台打包成单文件 exe（双击即用）
# 用法： bash scripts/build_windows.sh
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")/../web" && pwd)"   # <repo>/web
SK="$(cd "$HERE/.." && pwd)"
cd "$HERE"

PY="${PYTHON:-python}"

# 安装依赖（openpyxl / xlrd 为纯 Python；pyinstaller 仅打包期需要）
"$PY" -m pip install --quiet pyinstaller openpyxl xlrd || {
  echo "请先安装 Python 3 并加入 PATH，或设置 PYTHON 变量指向解释器。" >&2
  exit 1
}

# Windows 的 --add-data 分隔符是分号 ";"
"$PY" -m PyInstaller --onefile --name excel-a2b --noconsole \
  --add-data "$SK/web;web" --add-data "$SK/scripts;scripts" \
  --paths "$SK/scripts" --hidden-import openpyxl --hidden-import xlrd \
  serve.py

echo "完成：dist/excel-a2b.exe（双击运行，自动打开浏览器）。"
echo "分发时请把 exe 与说明一起提供；exe 输出目录下会生成转换结果文件。"
