"""客戶主檔對照：把辨識到的公司全稱 / 統編，對應到明細表使用的簡稱與統編。

資料來源為會計系統匯出的「客戶資料明細表」(.xls, BIFF)。
"""
from __future__ import annotations
import xlrd


def _tax_str(v) -> str:
    if isinstance(v, float):
        return str(int(v))
    return str(v).strip()


class CustomerMaster:
    def __init__(self, path: str):
        self.by_name: dict[str, tuple[str, str]] = {}   # 公司名稱 -> (簡稱, 統編)
        self.by_tax: dict[str, str] = {}                # 統編 -> 簡稱
        self.shorts: set[str] = set()
        self._load(path)

    def _load(self, path: str) -> None:
        wb = xlrd.open_workbook(path)
        for sh in wb.sheets():
            hdr = None
            for r in range(sh.nrows):
                row = [sh.cell_value(r, c) for c in range(sh.ncols)]
                if "公司名稱" in row and "簡稱" in row:
                    hdr = {v: i for i, v in enumerate(row)}
                    continue
                if not hdr:
                    continue
                name = str(row[hdr["公司名稱"]]).strip()
                short = str(row[hdr["簡稱"]]).strip()
                tax = _tax_str(row[hdr["統一編號"]])
                if name:
                    self.by_name[name] = (short, tax)
                if short:
                    self.shorts.add(short)
                if tax and short:
                    self.by_tax.setdefault(tax, short)

    def resolve(self, name: str | None, tax: str | None = None):
        """回傳 (簡稱, 統編, matched:bool)。

        優先以統編對應，其次公司全稱，再次簡稱本身即為主檔簡稱。
        """
        if tax:
            t = _tax_str(tax)
            if t in self.by_tax:
                return self.by_tax[t], t, True
        if name:
            nm = str(name).strip()
            if nm in self.by_name:
                short, t = self.by_name[nm]
                return short, t, True
            if nm in self.shorts:
                return nm, None, True
        return None, (_tax_str(tax) if tax else None), False
