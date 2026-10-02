#!/usr/bin/env python3
"""月資金計劃表 線上表單伺服器。

把資料夾裡「實際存在」的月資金計劃表.xlsx 變成瀏覽器可檢視、編輯、互動的表單；
每一次修改都寫回同一個原檔（自動、防抖 1 秒），並記錄到工作簿的「異動紀錄」分頁。

用法：
    python3 src/fund_plan_server.py "C:/Users/berti/OneDrive/文件/Claude/Projects/月資金計劃表.xlsx"
    python3 src/fund_plan_server.py                      # 依序找 FUND_PLAN_PATH 環境變數、fund_plan.config.json、data/*資金計劃*.xlsx
選項：
    --port 8765        監聽埠（預設 8765）
    --host 127.0.0.1   只允許本機瀏覽（預設）；區網共用改 0.0.0.0
    --open             啟動後自動開瀏覽器
    --no-log-sheet     不在工作簿內建立「異動紀錄」分頁（仍保留伺服器端記錄）
    --backup-dir DIR   備份資料夾（預設原檔同層的 備份/）

只用 Python 標準函式庫 + openpyxl，不需安裝 web 框架。
"""
from __future__ import annotations

import argparse
import glob
import json
import mimetypes
import os
import sys
import threading
import time
import traceback
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, unquote, urlparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fund_plan_model import EditError, ExternalModified, FundPlanBook  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
WEB_DIR = os.path.join(HERE, "fund_plan_web")
ROOT = os.path.dirname(HERE)
CONFIG_FILE = os.path.join(ROOT, "fund_plan.config.json")
SAVE_DEBOUNCE = 1.0      # 最後一次修改後 1 秒寫回
RETRY_INTERVAL = 5.0     # 寫回失敗（檔案被 Excel 鎖住）時每 5 秒重試


class AutoSaver(threading.Thread):
    """背景寫回：有變動就在防抖時間後存檔；失敗（檔案被鎖）則定期重試。"""

    def __init__(self, book: FundPlanBook):
        super().__init__(daemon=True, name="autosaver")
        self.book = book
        self.wake = threading.Event()
        self.stop = threading.Event()
        self.next_at = 0.0
        self.saving = False

    def schedule(self, delay: float = SAVE_DEBOUNCE) -> None:
        self.next_at = time.monotonic() + delay
        self.wake.set()

    def run(self) -> None:
        while not self.stop.is_set():
            self.wake.wait(timeout=1.0)
            self.wake.clear()
            if not self.book.dirty:
                continue
            now = time.monotonic()
            if now < self.next_at:
                continue
            self.saving = True
            try:
                self.book.save()
                print(f"[{time.strftime('%H:%M:%S')}] 已寫回 {self.book.path}")
            except ExternalModified as e:
                print("[衝突]", e)
                self.next_at = time.monotonic() + 3600  # 等使用者按重新載入
            except PermissionError:
                print("[等待] 原檔被鎖定，稍後重試")
                self.next_at = time.monotonic() + RETRY_INTERVAL
                self.wake.set()
            except Exception as e:  # pragma: no cover
                print("[錯誤] 寫回失敗:", e)
                traceback.print_exc()
                self.next_at = time.monotonic() + RETRY_INTERVAL
                self.wake.set()
            finally:
                self.saving = False


def find_default_path() -> str | None:
    env = os.environ.get("FUND_PLAN_PATH")
    if env and os.path.exists(env):
        return env
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, encoding="utf-8") as f:
                cfg = json.load(f)
            p = cfg.get("path")
            if p and os.path.exists(p):
                return p
        except (OSError, ValueError):
            pass
    hits = sorted(glob.glob(os.path.join(ROOT, "data", "*資金計劃*.xls[xm]")))
    return hits[0] if hits else None


def make_handler(book: FundPlanBook, saver: AutoSaver):
    class Handler(BaseHTTPRequestHandler):
        server_version = "FundPlanForm/1.0"

        def log_message(self, fmt, *args):  # 安靜一點，只印 API 錯誤
            if args and str(args[1]).startswith(("4", "5")):
                super().log_message(fmt, *args)

        # ---- 工具 ----
        def _json(self, obj, status=200):
            body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)

        def _body(self):
            n = int(self.headers.get("Content-Length") or 0)
            if not n:
                return {}
            return json.loads(self.rfile.read(n).decode("utf-8") or "{}")

        def _static(self, rel):
            path = os.path.normpath(os.path.join(WEB_DIR, rel.lstrip("/")))
            if not path.startswith(WEB_DIR) or not os.path.isfile(path):
                self.send_error(404)
                return
            ctype = mimetypes.guess_type(path)[0] or "application/octet-stream"
            if ctype.startswith("text/") or ctype in ("application/javascript", "application/json"):
                ctype += "; charset=utf-8"
            with open(path, "rb") as f:
                data = f.read()
            self.send_response(200)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(data)

        def _status(self):
            st = book.status()
            st["saving"] = saver.saving
            st["sync"] = ("conflict" if st["external_modified"] and st["dirty"] else
                          "error" if st["error"] else
                          "saving" if saver.saving else
                          "pending" if st["dirty"] else
                          "external" if st["external_modified"] else "clean")
            return st

        # ---- GET ----
        def do_GET(self):
            u = urlparse(self.path)
            path = unquote(u.path)
            q = parse_qs(u.query)
            try:
                if path in ("/", "/index.html"):
                    return self._static("index.html")
                if path == "/api/status":
                    return self._json(self._status())
                if path == "/api/workbook":
                    st = self._status()
                    st["sheets"] = book.sheet_names()
                    return self._json(st)
                if path.startswith("/api/sheet/"):
                    name = path[len("/api/sheet/"):]
                    if name not in book.sheets:
                        return self._json({"error": f"找不到分頁 {name}"}, 404)
                    payload = book.sheet_payload(name)
                    payload["status"] = self._status()
                    return self._json(payload)
                if path == "/api/log":
                    limit = int(q.get("limit", ["200"])[0])
                    return self._json({"changes": book.changes[-limit:][::-1]})
                if path.startswith("/static/"):
                    return self._static(path[len("/static/"):])
                return self.send_error(404)
            except Exception as e:  # pragma: no cover
                traceback.print_exc()
                return self._json({"error": str(e)}, 500)

        # ---- POST ----
        def do_POST(self):
            path = unquote(urlparse(self.path).path)
            try:
                data = self._body()
                if path == "/api/cell":
                    res = book.set_cell(data["sheet"], int(data["row"]), data["col"],
                                        data.get("value", ""), data.get("actor"))
                    if res.get("changed"):
                        saver.schedule()
                    res["sheet_data"] = book.sheet_payload(data["sheet"])
                    res["status"] = self._status()
                    return self._json(res)
                if path == "/api/cells":
                    results = []
                    changed = False
                    for it in data.get("edits", []):
                        try:
                            r = book.set_cell(it["sheet"], int(it["row"]), it["col"], it.get("value", ""),
                                              data.get("actor"))
                            changed = changed or r.get("changed", False)
                            results.append(r)
                        except EditError as e:
                            results.append({"changed": False, "row": it.get("row"), "col": it.get("col"),
                                            "error": str(e)})
                    if changed:
                        saver.schedule()
                    sheet = data.get("sheet") or (data["edits"][0]["sheet"] if data.get("edits") else None)
                    return self._json({"results": results,
                                       "sheet_data": book.sheet_payload(sheet) if sheet else None,
                                       "status": self._status()})
                if path == "/api/row/add":
                    r = book.add_row(data["sheet"], data.get("actor"))
                    saver.schedule()
                    return self._json({"row": r, "sheet_data": book.sheet_payload(data["sheet"]),
                                       "status": self._status()})
                if path == "/api/row/clear":
                    recs = book.clear_row(data["sheet"], int(data["row"]), data.get("actor"))
                    if recs:
                        saver.schedule()
                    return self._json({"cleared": len(recs), "sheet_data": book.sheet_payload(data["sheet"]),
                                       "status": self._status()})
                if path == "/api/save":
                    try:
                        res = book.save(force=bool(data.get("force")))
                    except ExternalModified as e:
                        return self._json({"error": str(e), "status": self._status()}, 409)
                    except PermissionError:
                        return self._json({"error": book.last_error, "status": self._status()}, 423)
                    res["status"] = self._status()
                    return self._json(res)
                if path == "/api/reload":
                    if book.dirty and not data.get("discard"):
                        return self._json({"error": "尚有未寫入的變動，確認要放棄再重新載入。",
                                           "status": self._status()}, 409)
                    t0 = time.time()
                    book.load()
                    return self._json({"reloaded": True, "seconds": round(time.time() - t0, 1),
                                       "status": self._status(), "sheets": book.sheet_names()})
                return self.send_error(404)
            except EditError as e:
                return self._json({"error": str(e), "status": self._status()}, 400)
            except (KeyError, ValueError) as e:
                return self._json({"error": f"參數錯誤：{e}"}, 400)
            except Exception as e:  # pragma: no cover
                traceback.print_exc()
                return self._json({"error": str(e)}, 500)

    return Handler


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="月資金計劃表線上表單")
    ap.add_argument("path", nargs="?", help="月資金計劃表 .xlsx 路徑（原檔，會直接寫回）")
    ap.add_argument("--port", type=int, default=int(os.environ.get("FUND_PLAN_PORT", "8765")))
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--open", action="store_true", help="啟動後自動開啟瀏覽器")
    ap.add_argument("--no-log-sheet", action="store_true")
    ap.add_argument("--backup-dir", default=None)
    a = ap.parse_args(argv)

    path = a.path or find_default_path()
    if not path or not os.path.exists(path):
        print("找不到月資金計劃表。請指定路徑：python3 src/fund_plan_server.py <月資金計劃表.xlsx>")
        print("或在 fund_plan.config.json 寫入 {\"path\": \"...\"}，或設定環境變數 FUND_PLAN_PATH。")
        return 2

    book = FundPlanBook(path, backup_dir=a.backup_dir, log_sheet=None if a.no_log_sheet else "異動紀錄")
    print(f"載入 {os.path.abspath(path)} …（巨大工作簿可能需要 2–5 分鐘）")
    t0 = time.time()
    book.load()
    print(f"載入完成 {time.time() - t0:.1f} 秒，可編輯分頁：{'、'.join(book.sheet_names()) or '（無）'}")
    if not book.sheets:
        print("[警告] 沒有偵測到含「日期／摘要」表頭的分頁，畫面會是空的。")

    saver = AutoSaver(book)
    saver.start()
    srv = ThreadingHTTPServer((a.host, a.port), make_handler(book, saver))
    url = f"http://{'localhost' if a.host in ('127.0.0.1', '0.0.0.0') else a.host}:{a.port}/"
    print(f"線上表單：{url}   （Ctrl+C 結束；結束前會把未寫入的變動寫回）")
    if a.open:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        saver.stop.set()
        if book.dirty:
            try:
                book.save()
                print("結束前已寫回未儲存的變動。")
            except Exception as e:
                print("[警告] 結束前寫回失敗：", e)
        srv.server_close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
