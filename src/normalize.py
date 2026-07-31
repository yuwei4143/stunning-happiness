"""欄位標準化規則：把支票辨識產出的原始欄位，轉成應付帳款明細表所用的格式。

所有規則皆由「使用者修正過的黃金樣板」(11505_updated) 反推而來，
集中在此檔以便日後維護與擴充。
"""
from __future__ import annotations
import datetime as _dt
import re

# --------------------------------------------------------------------------
# 銀行簡稱對照：辨識產出的銀行全稱 -> 明細表使用的簡稱
# 農會 / 漁會 / 信用部 命名不規則，另以 SPECIAL_ACCOUNT 處理或標記人工覆核。
# --------------------------------------------------------------------------
BANK_SHORT: dict[str, str] = {
    "中國信託銀行": "中信",
    "第一銀行": "一銀",
    "合作金庫商業銀行": "合庫",
    "彰化商業銀行": "彰銀",
    "臺灣中小企業銀行": "台企",
    "台中商業銀行": "台中銀",
    "華南商業銀行": "華南",
    "華泰商業銀行": "華泰",
    "玉山銀行": "玉山",
    "三信商業銀行": "三信銀",
    "永豐銀行": "永豐",
    "京城銀行": "京城",
    "台北富邦銀行": "台北富邦",
}

# 農漁會信用部等非商業銀行，分行命名不規則，逐筆對照（可持續擴充）。
# key = (銀行, 分行)；value = 明細表支票帳號的「機構+分行」前綴（帳號另接於後）。
SPECIAL_BRANCH: dict[tuple[str, str], str] = {
    ("秀水鄉農會", "農會"): "秀水農會",
    ("八里區農會", "龍形"): "八里農會",
}

# 完全特殊、無標準帳號者（例如彰化區漁會本會）。
SPECIAL_ACCOUNT: dict[tuple[str, str], str] = {
    ("彰化區漁會", "信用部"): "彰漁本會",
}


def clean_account(acct: str | None) -> str:
    """去除帳號中的破折號、空白等符號，只保留英數字。

    例：003269-8 -> 0032698、150-031-0018158-7 -> 15003100181587
    """
    if acct is None:
        return ""
    return re.sub(r"[^0-9A-Za-z]", "", str(acct))


def compose_check_account(bank: str | None, branch: str | None,
                          bank_branch: str | None, account: str | None):
    """組出明細表的「支票帳號」欄。

    回傳 (值, 是否需人工覆核, 覆核原因)。
    無法標準化的機構會回傳 needs_review=True，避免靜默填錯。
    """
    bank = (bank or "").strip()
    branch = (branch or "").strip()
    acct_clean = clean_account(account)

    # 1) 完全特殊機構（無標準帳號）
    if (bank, branch) in SPECIAL_ACCOUNT:
        return SPECIAL_ACCOUNT[(bank, branch)], False, ""

    # 2) 特殊分行前綴（農漁會）
    if (bank, branch) in SPECIAL_BRANCH:
        return SPECIAL_BRANCH[(bank, branch)] + acct_clean, False, ""

    # 3) 一般商業銀行
    if bank in BANK_SHORT:
        return f"{BANK_SHORT[bank]}{branch}{acct_clean}", False, ""

    # 4) 未知機構：以辨識的「銀行+分行」原樣拼帳號，但標記人工覆核
    fallback = f"{(bank_branch or (bank + branch)).strip()}{acct_clean}"
    return fallback, True, f"銀行簡稱未建檔（{bank}），支票帳號請人工確認"


def roc_to_date(value) -> _dt.date | None:
    """把日期正規化成 date。

    支援：民國格式字串 115/07/08、datetime、date。
    """
    if value is None or value == "":
        return None
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    s = str(value).strip()
    m = re.match(r"^(\d{2,3})/(\d{1,2})/(\d{1,2})", s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return _dt.date(y + 1911, mo, d)
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        return _dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def build_remark(memo: str | None, drawer: str | None, customer: str | None,
                 face_amount, invoice_total) -> str:
    """組出明細表「備註」欄。

    規則（由黃金樣板反推）：
      - 保留辨識備註中的「禁背」。
      - 發票人(drawer) 與客戶(customer) 不符時，附註發票人印章名稱。
      - 票面(face_amount) 與發票合計(invoice_total) 有差額時，附註待確認金額。
    """
    parts: list[str] = []
    memo = (memo or "").strip()
    if "禁背" in memo:
        parts.append("禁背")

    drawer = (drawer or "").strip()
    customer = (customer or "").strip()
    if drawer and customer and drawer != customer and drawer not in customer:
        # 個人印章 vs 公司印章的措辭差異保留給人工，這裡給出可辨識的基本註記
        parts.append(f"發票人印章:{drawer}")

    if face_amount is not None and invoice_total is not None:
        try:
            diff = int(invoice_total) - int(face_amount)
        except (TypeError, ValueError):
            diff = 0
        if diff > 0:
            parts.append(
                f"票面{int(face_amount):,}<合計{int(invoice_total):,} "
                f"差{diff:,}待確認"
            )
        elif diff < 0:
            parts.append(
                f"票面{int(face_amount):,}>合計{int(invoice_total):,} "
                f"差{-diff:,}待確認"
            )

    return " ".join(parts)
