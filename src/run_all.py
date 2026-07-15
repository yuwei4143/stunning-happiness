"""一鍵處理所有已知批次，輸出待回填明細、覆核彙總與名稱差異待建檔清單。"""
import os
import sys
sys.path.insert(0, os.path.dirname(__file__))

import openpyxl

from payable_filler import (KNOWN_BATCHES, process_batch, write_output,
                            load_master, CUSTOMER_ALIAS_PATH)

NAME_DIFF_REPORT = "output/名稱差異待建檔.xlsx"


def write_name_diff_report(rows: list[dict], path: str) -> None:
    """輸出名稱差異清單（去重）。欄位對齊別名對照檔，確認後可直接建檔。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "名稱差異待建檔"
    ws.append(["辨識名稱", "簡稱", "統編", "出現批次", "支票號碼"])
    for r in rows:
        ws.append([r["辨識名稱"], "", "",
                   "、".join(sorted(r["批次"])), "、".join(sorted(r["票號"]))])
    wb.save(path)


def main() -> None:
    os.makedirs("output", exist_ok=True)
    master = load_master()
    if master is None:
        print("[提醒] 找不到客戶主檔，客戶簡稱/統編將留白並標記人工。")

    summary = []
    name_diffs: dict[str, dict] = {}   # 辨識名稱 -> 彙整
    for path, batch in KNOWN_BATCHES.items():
        name = os.path.splitext(os.path.basename(path))[0]
        filled = process_batch(path, batch, master=master)
        out = f"output/{name}_待回填.xlsx"
        write_output(filled, out)
        n_rv = sum(1 for f in filled if f.need_review)
        summary.append((name, len(filled), n_rv, out))
        for fc in filled:
            if fc.name_mismatch and fc.客戶:
                d = name_diffs.setdefault(
                    fc.客戶, {"辨識名稱": fc.客戶, "批次": set(), "票號": set()})
                d["批次"].add(name)
                d["票號"].add(fc.支票號碼)

    print("批次處理完成：")
    for name, n, rv, out in summary:
        print(f"  {name}: {n} 筆，{rv} 筆需人工覆核 → {out}")

    if name_diffs:
        write_name_diff_report(list(name_diffs.values()), NAME_DIFF_REPORT)
        print(f"\n名稱差異 {len(name_diffs)} 筆 → {NAME_DIFF_REPORT}")
        print(f"  確認簡稱/統編後，填入 {CUSTOMER_ALIAS_PATH}，往後即自動解析。")
        for nm in name_diffs:
            print(f"    - {nm}")
    else:
        print("\n無名稱差異，全部客戶皆已對應主檔。")


if __name__ == "__main__":
    main()
