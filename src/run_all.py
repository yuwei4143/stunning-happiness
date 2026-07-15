"""一鍵處理所有已知批次，輸出待回填明細與覆核彙總。"""
import os
import sys
sys.path.insert(0, os.path.dirname(__file__))

from payable_filler import (KNOWN_BATCHES, process_batch, write_output,
                            load_master)


def main() -> None:
    os.makedirs("output", exist_ok=True)
    master = load_master()
    if master is None:
        print("[提醒] 找不到客戶主檔，客戶簡稱/統編將留白並標記人工。")
    summary = []
    for path, batch in KNOWN_BATCHES.items():
        name = os.path.splitext(os.path.basename(path))[0]
        filled = process_batch(path, batch, master=master)
        out = f"output/{name}_待回填.xlsx"
        write_output(filled, out)
        n_rv = sum(1 for f in filled if f.need_review)
        summary.append((name, len(filled), n_rv, out))

    print("批次處理完成：")
    for name, n, rv, out in summary:
        print(f"  {name}: {n} 筆，{rv} 筆需人工覆核 → {out}")


if __name__ == "__main__":
    main()
