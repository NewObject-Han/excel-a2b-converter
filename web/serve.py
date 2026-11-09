#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
serve.py — Excel A→B 操作台的本地引擎服务

作用
----
1. 把 web/ 目录作为静态站点提供（index.html + assets + vendor）。
2. 暴露本地引擎接口，让操作页能用 Python 引擎执行转换（完整保留模板格式）：
   GET  /api/health   -> {"ok": true, ...}
   POST /api/convert  -> 接收 JSON（文件 base64 + mapping），返回结果 xlsx 的 base64

用法
----
  python serve.py                 # 默认 127.0.0.1:8765，自动打开浏览器
  python serve.py --port 9000
  python serve.py --no-browser

注意：需用装好依赖的解释器运行（见 SKILL.md，可先执行 scripts/bootstrap.sh）。
"""
from __future__ import annotations

import argparse
import base64
import json
import os
import sys
import tempfile
import threading
import webbrowser
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

HERE = os.path.dirname(os.path.abspath(__file__))

# 同时支持「源码运行」与「PyInstaller 冻结后运行」两种形态：
# 冻结时资源被解压到 sys._MEIPASS 下，目录结构与仓库一致（web/、scripts/）。
if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    _BASE = sys._MEIPASS
    WEB_DIR = os.path.join(_BASE, "web")
    SCRIPTS_DIR = os.path.join(_BASE, "scripts")
else:
    WEB_DIR = HERE
    SCRIPTS_DIR = os.path.normpath(os.path.join(HERE, "..", "scripts"))

sys.path.insert(0, SCRIPTS_DIR)


def _check_deps():
    missing = []
    try:
        import openpyxl  # noqa: F401
    except Exception:
        missing.append("openpyxl")
    try:
        import xlrd  # noqa: F401
    except Exception:
        missing.append("xlrd")
    if missing:
        print("缺少依赖: " + ", ".join(missing), file=sys.stderr)
        print("请先运行: bash scripts/bootstrap.sh", file=sys.stderr)
        print("或: python -m pip install openpyxl xlrd", file=sys.stderr)
        sys.exit(2)


MAX_BODY = 200 * 1024 * 1024  # 200MB


def _ext(name: str, fallback: str = ".xlsx") -> str:
    e = os.path.splitext(name or "")[1].lower()
    return e if e in (".xls", ".xlsx", ".xlsm", ".csv") else fallback


def do_convert(payload: dict) -> dict:
    """执行一次转换，返回 {ok, filename, b64, stats} 或 {ok:false, error}"""
    import excel_a2b

    tmpdir = tempfile.mkdtemp(prefix="excel_a2b_")
    try:
        src = payload.get("source") or {}
        if not src.get("b64"):
            raise ValueError("缺少源表文件")
        src_path = os.path.join(tmpdir, "source" + _ext(src.get("name")))
        with open(src_path, "wb") as f:
            f.write(base64.b64decode(src["b64"]))

        tgt = payload.get("target") or None
        tpl_path = None
        if tgt and tgt.get("b64"):
            tpl_path = os.path.join(tmpdir, "template" + _ext(tgt.get("name")))
            with open(tpl_path, "wb") as f:
                f.write(base64.b64decode(tgt["b64"]))

        out_name = payload.get("out_name") or "转换结果.xlsx"
        base = os.path.basename(out_name)
        # 输出落点：尽量放在程序所在目录（便于用户找到），不可写则退回临时目录
        if getattr(sys, "frozen", False):
            out_path = os.path.join(os.path.dirname(sys.executable), base)
        else:
            out_path = os.path.join(tmpdir, base)
        try:
            open(out_path, "ab").close()
        except Exception:
            import tempfile as _tf
            out_path = os.path.join(_tf.gettempdir(), base)

        cfg = {
            "source": {
                "path": src_path,
                "sheet": src.get("sheet"),
                "header_row": src.get("header_row") or 1,
            },
            "target": {
                "template": tpl_path,
                "path": out_path,
                "sheet": (tgt or {}).get("sheet"),
                "header_row": (tgt or {}).get("header_row") or 1,
            },
            "mapping": payload.get("mapping") or [],
            "options": payload.get("options") or {},
        }
        stats = excel_a2b.convert(cfg)
        with open(out_path, "rb") as f:
            data = f.read()
        return {
            "ok": True,
            "filename": os.path.basename(out_path),
            "saved_to": out_path,
            "b64": base64.b64encode(data).decode("ascii"),
            "stats": stats,
        }
    finally:
        # 清理临时输入文件（保留 out_path，它可能落盘在桌面）
        try:
            import shutil
            for fn in os.listdir(tmpdir):
                p = os.path.join(tmpdir, fn)
                if os.path.abspath(p) != os.path.abspath(locals().get("out_path", "")):
                    try:
                        os.remove(p)
                    except Exception:
                        pass
            shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception:
            pass


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=WEB_DIR, **kw)

    # 静默日志噪音
    def log_message(self, fmt, *args):
        if "/api/" in (self.path or ""):
            sys.stderr.write("[api] %s\n" % (fmt % args))

    def _json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path.split("?")[0] == "/api/health":
            return self._json({"ok": True, "engine": "python", "version": 1})
        return super().do_GET()

    def do_POST(self):
        if self.path.split("?")[0] != "/api/convert":
            return self._json({"ok": False, "error": "未知接口"}, 404)
        try:
            length = int(self.headers.get("Content-Length") or 0)
            if length <= 0 or length > MAX_BODY:
                raise ValueError("请求体为空或过大")
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            res = do_convert(payload)
            return self._json(res)
        except Exception as e:
            import traceback
            traceback.print_exc()
            return self._json({"ok": False, "error": "%s: %s" % (type(e).__name__, e)}, 200)

    def end_headers(self):
        # 便于开发时看到最新静态文件
        if self.path and not self.path.startswith("/api/"):
            self.send_header("Cache-Control", "no-cache")
        super().end_headers()


def main():
    ap = argparse.ArgumentParser(description="Excel A→B 操作台本地服务")
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--no-browser", action="store_true")
    args = ap.parse_args()

    _check_deps()

    server = None
    port = args.port
    for _ in range(20):
        try:
            server = ThreadingHTTPServer((args.host, port), Handler)
            break
        except OSError:
            port += 1
    if server is None:
        print("无法绑定端口", file=sys.stderr)
        sys.exit(1)

    url = "http://%s:%d/" % (args.host, port)
    print("=" * 58)
    print("  Excel A→B 转换操作台")
    print("  地址: " + url)
    print("  引擎: Python（保留模板格式）")
    print("  按 Ctrl+C 停止")
    print("=" * 58)

    if not args.no_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
