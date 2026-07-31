"""支票辨識產出 → 應付帳款明細表 的自動預填工具。

流程：
  1. 讀取支票辨識紀錄（辨識專案產出的 xlsx）。
  2. 依 normalize.py 的規則，把每筆支票轉成明細表所用欄位格式。
  3. 套用批次層級的覆寫（用途、送票銀行/日/帳號）。
  4. 輸出「待回填明細」：與明細表相同欄位排列，供覆核後貼入主檔。
  5. 需人工判斷者（辨識覆核旗標、銀行未建檔、金額待確認）一律標記，不靜默填錯。

設計原則：能由規則確定的自動填；需判斷的標記 need_review，交由使用者最後確認。
"""
from __future__ import annotations
import argparse
import datetime as _dt
from dataclasses import dataclass, field, asdict

import openpyxl

from customer_master import CustomerMaster
from normalize import (
    compose_check_account,
    roc_to_date,
    build_remark,
)

# 送票帳號 -> 用途 對照（由黃金樣板反推：不同託收/交換帳戶對應不同用途）。
COLLECTION_ACCOUNT_PURPOSE: dict[str, str] = {
    "15000100182967": "託收",
    "57100101404800": "交換",
}

# 明細表支票區欄位順序（對應主檔第 10~19 欄）。
CHECK_COLUMNS = [
    "支票帳號", "支票號碼", "到期日", "備註", "收票日",
    "實收金額", "送票銀行", "送票日", "送票帳號", "用途",
]


@dataclass
class BatchSettings:
    """一個辨識批次的送票資訊；缺漏欄位會以辨識紀錄逐筆值補上。"""
    送票銀行: str | None = None
    送票日: _dt.date | None = None
    送票帳號: str | None = None
    用途: str | None = None  # 交換 / 託收；None 表示依送票帳號自動判斷


@dataclass
class FilledCheck:
    支票帳號: str = ""
    支票號碼: str = ""
    到期日: _dt.date | None = None
    備註: str = ""
    收票日: _dt.date | None = None
    實收金額: int | None = None
    送票銀行: str | None = None
    送票日: _dt.date | None = None
    送票帳號: str | None = None
    用途: str | None = None
    # 輔助 / 覆核欄位
    客戶: str = ""
    客戶簡稱: str = ""      # 由客戶主檔對應（明細表所用簡稱）
    統編: str | None = None
    name_mismatch: bool = False   # 辨識客戶在主檔查無（名稱差異，另行註記）
    need_review: bool = False
    review_reasons: list[str] = field(default_factory=list)

    def flag(self, reason: str) -> None:
        self.need_review = True
        if reason and reason not in self.review_reasons:
            self.review_reasons.append(reason)


def load_check_records(path: str) -> list[dict]:
    """讀取支票辨識紀錄工作表，回傳每列的欄名->值 dict。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["支票辨識紀錄"]
    header = [c.value for c in ws[1]]
    records = []
    for r in range(2, ws.max_row + 1):
        values = [ws.cell(r, c).value for c in range(1, len(header) + 1)]
        if all(v is None or str(v).strip() == "" for v in values):
            continue
        if not values[header.index("支票號碼")]:
            continue
        records.append(dict(zip(header, values)))
    return records


def _to_bool(v) -> bool:
    return str(v).strip().upper() in ("TRUE", "1", "是", "Y", "YES")


def fill_record(rec: dict, batch: BatchSettings,
                invoice_total: int | None = None,
                master=None) -> FilledCheck:
    """把單筆辨識紀錄轉成明細表格式的 FilledCheck。

    invoice_total：若已知對應發票合計，用於金額勾稽備註（可為 None）。
    master：CustomerMaster 實例，用於對應客戶簡稱與統編（可為 None）。
    """
    out = FilledCheck()

    # 送票資訊：優先用辨識紀錄，缺漏才用批次設定。
    out.送票銀行 = (rec.get("送票銀行") or batch.送票銀行) or None
    out.送票帳號 = str(rec.get("送票號碼") or batch.送票帳號 or "").strip() or None
    out.送票日 = roc_to_date(rec.get("送票日")) or batch.送票日

    # 用途：批次覆寫 > 送票帳號推斷 > 辨識逐筆值。
    if batch.用途:
        out.用途 = batch.用途
    elif out.送票帳號 in COLLECTION_ACCOUNT_PURPOSE:
        out.用途 = COLLECTION_ACCOUNT_PURPOSE[out.送票帳號]
    else:
        out.用途 = (rec.get("用途") or None)

    # 支票帳號（銀行簡稱標準化 + 帳號去符號）。
    acct, need_rv, reason = compose_check_account(
        rec.get("銀行"), rec.get("分行"),
        rec.get("銀行+分行"), rec.get("支存帳號"),
    )
    out.支票帳號 = acct
    if need_rv:
        out.flag(reason)

    out.支票號碼 = str(rec.get("支票號碼") or "").strip()
    out.到期日 = roc_to_date(rec.get("到期日"))
    out.收票日 = roc_to_date(rec.get("收票日"))
    out.客戶 = str(rec.get("客戶") or "").strip()

    face = rec.get("實收金額")
    face = int(face) if isinstance(face, (int, float)) else None

    # 實收金額：僅「交換」用途填入票面金額；「託收」留空（尚未兌現）。
    out.實收金額 = face if out.用途 == "交換" else None

    # 備註：禁背 + 發票人印章不符 + 金額勾稽差額。
    out.備註 = build_remark(
        rec.get("備註"), rec.get("發票人"), rec.get("客戶"),
        face, invoice_total,
    )

    # 客戶主檔對應：解出明細表所用簡稱與統編。
    if master is not None:
        short, tax, matched = master.resolve(rec.get("客戶"))
        if matched:
            out.客戶簡稱 = short or ""
            out.統編 = tax
        else:
            out.name_mismatch = True
            out.flag(f"【名稱差異】辨識客戶「{out.客戶}」主檔查無，"
                     f"請建檔對照（簡稱/統編待補）")

    # 覆核旗標：辨識端已標記者一律帶入。
    if _to_bool(rec.get("需人工覆核")):
        out.flag(str(rec.get("覆核原因") or "辨識端標記需人工覆核").strip())

    return out


def process_batch(path: str, batch: BatchSettings, master=None) -> list[FilledCheck]:
    filled = [fill_record(rec, batch, master=master)
              for rec in load_check_records(path)]
    _flag_combine_candidates(filled)
    return filled


def _flag_combine_candidates(filled: list[FilledCheck]) -> None:
    """同一客戶＋同一支存帳號出現多張票 → 疑似需合併至同一發票列，標記人工。

    合併時的「實收金額（加總）」與「到期日（取捨）」需人工決定，故不自動填死。
    """
    from collections import defaultdict
    groups: dict[tuple[str, str], list[FilledCheck]] = defaultdict(list)
    for fc in filled:
        groups[(fc.客戶, fc.支票帳號)].append(fc)
    for (cust, _), members in groups.items():
        if len(members) > 1:
            nums = "、".join(m.支票號碼 for m in members)
            for m in members:
                m.flag(f"同客戶多張票（{nums}）疑似合併，實收/到期請人工確認")


def write_output(filled: list[FilledCheck], out_path: str) -> None:
    """輸出待回填明細（含覆核欄）到 xlsx。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "待回填明細"
    headers = (["統編", "客戶簡稱"] + CHECK_COLUMNS
               + ["辨識客戶", "需人工覆核", "覆核原因"])
    ws.append(headers)
    for fc in filled:
        d = asdict(fc)
        row = [fc.統編, fc.客戶簡稱] + [d[c] for c in CHECK_COLUMNS]
        row += [fc.客戶, "是" if fc.need_review else "",
                "；".join(fc.review_reasons)]
        ws.append(row)
    wb.save(out_path)


# --------------------------------------------------------------------------
# 三個已知批次的設定（可依實際送票資訊調整）。
# --------------------------------------------------------------------------
KNOWN_BATCHES: dict[str, BatchSettings] = {
    "data/checks_20260618.xlsx": BatchSettings(
        送票銀行="彰銀", 送票帳號="57100101404800",
        送票日=_dt.date(2026, 6, 22), 用途="交換"),
    "data/checks_20260624.xlsx": BatchSettings(
        送票銀行="彰銀", 送票帳號="57100101404800",
        送票日=None, 用途="交換"),
    "data/checks_20260708.xlsx": BatchSettings(
        送票銀行="永豐", 送票帳號="15000100182967",
        送票日=_dt.date(2026, 7, 8), 用途="託收"),
}

# 客戶主檔路徑（會計系統匯出的客戶資料明細表）。
CUSTOMER_MASTER_PATH = "data/customer_master.xls"
# 名稱差異別名對照（使用者建檔後往後自動解析）。
CUSTOMER_ALIAS_PATH = "data/customer_alias.csv"


def load_master(path: str = CUSTOMER_MASTER_PATH,
                alias_path: str = CUSTOMER_ALIAS_PATH):
    try:
        return CustomerMaster(path, alias_path=alias_path)
    except FileNotFoundError:
        return None


def main() -> None:
    ap = argparse.ArgumentParser(description="支票辨識產出 → 應付帳款明細預填")
    ap.add_argument("check_file", help="支票辨識紀錄 xlsx 路徑")
    ap.add_argument("-o", "--output", default="output/filled.xlsx")
    ap.add_argument("--purpose", choices=["交換", "託收"], help="批次用途覆寫")
    ap.add_argument("--send-bank")
    ap.add_argument("--send-account")
    args = ap.parse_args()

    batch = KNOWN_BATCHES.get(
        args.check_file,
        BatchSettings(送票銀行=args.send_bank, 送票帳號=args.send_account,
                      用途=args.purpose),
    )
    if args.purpose:
        batch.用途 = args.purpose

    filled = process_batch(args.check_file, batch, master=load_master())
    write_output(filled, args.output)
    n_rv = sum(1 for f in filled if f.need_review)
    print(f"已處理 {len(filled)} 筆，其中 {n_rv} 筆需人工覆核 → {args.output}")


if __name__ == "__main__":
    main()
