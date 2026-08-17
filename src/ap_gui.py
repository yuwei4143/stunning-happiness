"""應付帳款自動預填 — 拖放式圖形介面。

給不熟指令列的同事：拖入（或選擇）支票辨識紀錄檔 → 按開始 → 產出待回填明細。
可用 PyInstaller 打包成單一執行檔，雙擊即用（見 packaging/ 說明）。

拖放需要 tkinterdnd2；若未安裝，仍可用「選擇檔案」按鈕，功能不受影響。
"""
from __future__ import annotations
import os
import sys
import threading
import datetime as _dt

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# 讓 PyInstaller onefile 與原始碼執行都能找到同層模組
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from payable_filler import (BatchSettings, process_batch, write_output,
                            COLLECTION_ACCOUNT_PURPOSE)
from customer_master import CustomerMaster

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _HAS_DND = True
except Exception:
    _HAS_DND = False

APP_TITLE = "應付帳款自動預填工具"


def _resource_dir() -> str:
    """回傳資源目錄（PyInstaller onefile 解壓目錄或原始碼目錄）。"""
    return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))


def _guess_master() -> str:
    """自動尋找客戶主檔：優先執行檔旁、其次專案 data/、再其次打包資源。"""
    here = os.path.dirname(os.path.abspath(sys.argv[0]))
    for cand in (
        os.path.join(here, "customer_master.xls"),
        os.path.join(here, "data", "customer_master.xls"),
        os.path.join(os.getcwd(), "data", "customer_master.xls"),
        os.path.join(_resource_dir(), "customer_master.xls"),
    ):
        if os.path.exists(cand):
            return cand
    return ""


class App:
    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        root.geometry("620x560")
        root.minsize(560, 520)

        self.check_path = tk.StringVar()
        self.master_path = tk.StringVar(value=_guess_master())
        self.out_dir = tk.StringVar()
        self.purpose = tk.StringVar(value="自動")
        self.send_bank = tk.StringVar()
        self.send_acct = tk.StringVar()
        self.send_date = tk.StringVar()

        self._build()

    # ---- UI ----
    def _build(self):
        pad = {"padx": 14, "pady": 6}
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        head = ttk.Frame(self.root)
        head.pack(fill="x", **pad)
        ttk.Label(head, text=APP_TITLE, font=("", 16, "bold")).pack(anchor="w")
        ttk.Label(head, text="拖入支票辨識紀錄檔，或按「選擇檔案」。",
                  foreground="#666").pack(anchor="w")

        # 拖放區
        self.drop = tk.Label(
            self.root, text="⬇  將支票辨識紀錄 .xlsx 拖到這裡  ⬇"
            if _HAS_DND else "請按下方「選擇檔案」",
            relief="ridge", bd=2, height=4, bg="#f3efe6", fg="#8a2822",
            font=("", 12, "bold"))
        self.drop.pack(fill="x", padx=14, pady=(4, 8))
        if _HAS_DND:
            self.drop.drop_target_register(DND_FILES)
            self.drop.dnd_bind("<<Drop>>", self._on_drop)

        self._file_row("支票辨識紀錄（必填）", self.check_path,
                       self._pick_check, filetypes=[("Excel", "*.xlsx")])
        self._file_row("客戶主檔（選填，自動偵測）", self.master_path,
                       self._pick_master,
                       filetypes=[("Excel", "*.xls *.xlsx")])

        # 批次設定
        box = ttk.LabelFrame(self.root, text="批次設定")
        box.pack(fill="x", padx=14, pady=8)
        r = ttk.Frame(box); r.pack(fill="x", padx=10, pady=6)
        ttk.Label(r, text="用途：").pack(side="left")
        for v in ("自動", "交換", "託收"):
            ttk.Radiobutton(r, text=v, value=v,
                            variable=self.purpose).pack(side="left", padx=4)
        ttk.Label(box, text="（自動＝依送票帳號判斷；查無則留用辨識值）",
                  foreground="#888").pack(anchor="w", padx=12)
        r2 = ttk.Frame(box); r2.pack(fill="x", padx=10, pady=6)
        ttk.Label(r2, text="送票銀行", width=8).grid(row=0, column=0, sticky="w")
        ttk.Entry(r2, textvariable=self.send_bank, width=12).grid(row=0, column=1, padx=4)
        ttk.Label(r2, text="送票帳號", width=8).grid(row=0, column=2, sticky="w")
        ttk.Entry(r2, textvariable=self.send_acct, width=18).grid(row=0, column=3, padx=4)
        r3 = ttk.Frame(box); r3.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Label(r3, text="送票日", width=8).grid(row=0, column=0, sticky="w")
        ttk.Entry(r3, textvariable=self.send_date, width=14).grid(row=0, column=1, padx=4)
        ttk.Label(r3, text="格式 2026-07-08（選填）",
                  foreground="#888").grid(row=0, column=2, sticky="w", padx=6)

        self._file_row("輸出資料夾（預設同來源檔）", self.out_dir,
                       self._pick_outdir, is_dir=True)

        # 動作
        act = ttk.Frame(self.root); act.pack(fill="x", padx=14, pady=8)
        self.run_btn = ttk.Button(act, text="開始產出", command=self._run)
        self.run_btn.pack(side="left")
        ttk.Button(act, text="離開", command=self.root.destroy).pack(side="right")

        # 記錄
        self.log = tk.Text(self.root, height=7, wrap="word", bg="#faf8f3")
        self.log.pack(fill="both", expand=True, padx=14, pady=(4, 12))
        self._say("準備就緒。" + ("（拖放已啟用）" if _HAS_DND else
                                  "（未偵測到拖放套件，請用選擇檔案）"))

    def _file_row(self, label, var, cmd, filetypes=None, is_dir=False):
        f = ttk.Frame(self.root); f.pack(fill="x", padx=14, pady=2)
        ttk.Label(f, text=label).pack(anchor="w")
        g = ttk.Frame(f); g.pack(fill="x")
        ttk.Entry(g, textvariable=var).pack(side="left", fill="x", expand=True)
        ttk.Button(g, text="選擇資料夾" if is_dir else "選擇檔案",
                   command=cmd).pack(side="left", padx=(6, 0))

    # ---- 事件 ----
    def _on_drop(self, event):
        path = event.data.strip().strip("{}")
        if path.lower().endswith(".xlsx"):
            self.check_path.set(path)
            self._say(f"已載入：{os.path.basename(path)}")
        else:
            self._say("請拖入 .xlsx 的支票辨識紀錄檔。")

    def _pick_check(self):
        p = filedialog.askopenfilename(title="選擇支票辨識紀錄",
                                       filetypes=[("Excel", "*.xlsx")])
        if p:
            self.check_path.set(p)

    def _pick_master(self):
        p = filedialog.askopenfilename(title="選擇客戶主檔",
                                       filetypes=[("Excel", "*.xls *.xlsx")])
        if p:
            self.master_path.set(p)

    def _pick_outdir(self):
        p = filedialog.askdirectory(title="選擇輸出資料夾")
        if p:
            self.out_dir.set(p)

    def _say(self, msg):
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.root.update_idletasks()

    # ---- 執行 ----
    def _run(self):
        chk = self.check_path.get().strip()
        if not chk or not os.path.exists(chk):
            messagebox.showwarning(APP_TITLE, "請先選擇支票辨識紀錄檔。")
            return
        self.run_btn.config(state="disabled")
        threading.Thread(target=self._work, args=(chk,), daemon=True).start()

    def _work(self, chk):
        try:
            send_date = None
            ds = self.send_date.get().strip()
            if ds:
                try:
                    send_date = _dt.date.fromisoformat(ds)
                except ValueError:
                    self._say(f"送票日格式無法解析（{ds}），將略過。")
            purpose = None if self.purpose.get() == "自動" else self.purpose.get()
            batch = BatchSettings(
                送票銀行=self.send_bank.get().strip() or None,
                送票帳號=self.send_acct.get().strip() or None,
                送票日=send_date, 用途=purpose)

            master = None
            mp = self.master_path.get().strip()
            if mp and os.path.exists(mp):
                try:
                    master = CustomerMaster(mp)
                    self._say(f"已載入客戶主檔：{os.path.basename(mp)}")
                except Exception as e:
                    self._say(f"客戶主檔載入失敗（{e}），客戶簡稱將留白。")
            else:
                self._say("未提供客戶主檔，客戶簡稱／統編將留白並標記人工。")

            self._say("處理中…")
            filled = process_batch(chk, batch, master=master)

            out_dir = self.out_dir.get().strip() or os.path.dirname(chk)
            base = os.path.splitext(os.path.basename(chk))[0]
            out_path = os.path.join(out_dir, f"{base}_待回填.xlsx")
            write_output(filled, out_path)

            n = len(filled)
            rv = sum(1 for f in filled if f.need_review)
            self._say(f"完成：{n} 筆，其中 {rv} 筆需人工覆核。")
            self._say(f"輸出 → {out_path}")
            self._done(out_path, n, rv, out_dir)
        except Exception as e:
            self._say(f"發生錯誤：{e}")
            self.root.after(0, lambda: messagebox.showerror(APP_TITLE, str(e)))
        finally:
            self.root.after(0, lambda: self.run_btn.config(state="normal"))

    def _done(self, out_path, n, rv, out_dir):
        def show():
            if messagebox.askyesno(
                    APP_TITLE,
                    f"完成！{n} 筆，{rv} 筆需人工覆核。\n\n是否開啟輸出資料夾？"):
                _open_folder(out_dir)
        self.root.after(0, show)


def _open_folder(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa
        elif sys.platform == "darwin":
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')
    except Exception:
        pass


def main():
    root = TkinterDnD.Tk() if _HAS_DND else tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
