# 應付帳款自動預填工具

> 圖文操作手冊：`docs/manual.html`（瀏覽器開啟）｜文字版：`docs/使用手冊.md`
>
> **月資金計劃表線上表單**（瀏覽器檢視／編輯，變動即時寫回原檔）：見下方「月資金計劃表 線上表單」。

把「支票辨識專案」的產出，依**使用者修正過的樣板**規則自動整理成
應付帳款明細表格式，供覆核後貼入主檔。

## 這個工具做什麼

支票辨識產出的原始欄位，離「填進明細表」還差一層加工。本工具把其中
**可由規則確定的部分自動完成**，把**需人工判斷的部分明確標記**，
不靜默填錯。

### 自動完成（固定規則，對黃金樣板驗證 100% 相符）

| 明細表欄位 | 規則 |
|-----------|------|
| 支票帳號 | 銀行簡稱標準化（中國信託→中信、第一銀行→一銀、合作金庫→合庫、彰化商銀→彰銀、臺灣中小企銀→台企、台中商銀→台中銀、三信→三信銀…）＋ 支存帳號去符號（003269-8→0032698） |
| 到期日／收票日／送票日 | 民國日期、datetime 一律正規化為日期 |
| 用途 | 依批次覆寫（送票帳號 15000100182967→託收、57100101404800→交換） |
| 實收金額 | 僅「交換」批填入票面；「託收」批留空 |
| 送票銀行／送票帳號 | 依辨識紀錄，缺漏用批次設定補 |
| 備註 | 保留「禁背」 |
| 客戶簡稱／統編 | 由客戶主檔（客戶資料明細表）以統編、公司全稱對應 |

### 標記人工覆核（需判斷，工具不臆測）

- 辨識端已標記 `需人工覆核=TRUE` 者，帶入原覆核原因。
- **合併票**：同客戶同帳號多張票，實收加總與到期取捨請人工確認。
- **金額勾稽**：票面與發票合計有差額（需比對主檔發票列）。
- **名稱差異（另行註記）**：辨識客戶在主檔查無時，以 `【名稱差異】` 專屬註記，
  並彙整到 `output/名稱差異待建檔.xlsx`（去重）。確認簡稱/統編後填入
  `data/customer_alias.csv`，往後同名稱自動解析、不再列入。
- 銀行簡稱未建檔的機構。

### 名稱差異建檔流程

1. 執行 `python3 src/run_all.py`，查看 `output/名稱差異待建檔.xlsx`。
2. 逐筆確認正確的「簡稱／統編」。
3. 把辨識名稱、簡稱、統編填入 `data/customer_alias.csv`。
4. 下次執行即自動對應，該名稱不再出現在名稱差異清單。

## 使用方式

```bash
pip install openpyxl

# 處理單一批次
python3 src/payable_filler.py data/checks_20260708.xlsx -o output/filled.xlsx

# 一鍵處理所有已知批次
python3 src/run_all.py

# 對使用者修正版樣板做黃金驗證
python3 tests/verify_against_golden.py
```

輸出為「待回填明細」工作表，欄位順序與明細表支票區一致，並附
`需人工覆核`／`覆核原因` 兩欄。覆核完成後即可貼入 11505 應付帳款明細表。

## 檔案結構

```
src/normalize.py            欄位標準化規則（銀行簡稱、帳號、日期、備註）
src/customer_master.py      客戶主檔對照（全稱/統編 → 簡稱）＋名稱差異別名檔
data/customer_alias.csv     名稱差異別名對照（使用者逐步建檔）
src/payable_filler.py       核心：辨識紀錄 → 待回填明細；批次設定
src/run_all.py              一鍵處理所有批次
src/fund_plan_server.py     月資金計劃表線上表單伺服器（標準函式庫 http.server）
src/fund_plan_model.py      資金計劃表讀取／編輯／寫回原檔／異動紀錄
src/formula_eval.py         極簡 Excel 公式評估器（公式欄即時顯示）
src/fund_plan_web/index.html 線上表單前端
tests/verify_against_golden.py  對 11505_updated 逐欄驗證
tests/test_fund_plan.py     線上表單自動化測試
tests/make_fund_plan_sample.py  產生同構的月資金計劃表範例
data/                       樣本：辨識批次、發票底稿、修正版樣板
output/                     產出的待回填明細
```

## 維護

新增銀行簡稱：編輯 `src/normalize.py` 的 `BANK_SHORT`。
新增農漁會分行：`SPECIAL_BRANCH` / `SPECIAL_ACCOUNT`。
新增批次送票設定：`src/payable_filler.py` 的 `KNOWN_BATCHES`。

---

## 月資金計劃表 線上表單

把資料夾裡**實際存在**的 `月資金計劃表.xlsx` 變成瀏覽器可檢視、編輯、互動的表單，
每一格的修改都**直接寫回同一個原檔**（不另存副本），並記錄在工作簿的「異動紀錄」分頁。

```bash
pip install -r requirements.txt
python3 src/fund_plan_server.py "C:/Users/berti/OneDrive/文件/Claude/Projects/月資金計劃表.xlsx" --open
# Windows：把 xlsx 拖到 啟動資金計劃表.bat 上，或先把路徑寫進 fund_plan.config.json（見 .example）
```

開 `http://localhost:8765/`：

| 功能 | 說明 |
|------|------|
| 月份分頁 | 自動列出所有 `2026(N)` 分頁；`YYYYMMDD` 版本分頁與「異動紀錄」不列入 |
| 可編輯欄 | 日期／摘要／支出／提前扣款／未入餘額／備註等**非公式欄**；離開格子即寫回 |
| 公式欄（灰底 ƒ） | 實際應付／應付餘額／目前現金等 Excel 表格 calculated column，**唯讀**、由內建公式評估器即時算出顯示；存檔時保留原公式，Excel 開檔全量重算 |
| KPI | 支出合計、收入合計、月底目前現金、最低現金點、低於警戒線的天數 |
| 互動 | 關鍵字篩選、隱藏空白列、現金警戒線標色、`Enter` 往下／`Tab` 往右／`Esc` 還原、新增列（延伸表格並複製公式）、清空列 |
| 同步狀態 | 已同步／待寫回／寫入中／衝突；原檔被 Excel 鎖住會自動重試；原檔被外部另存會停止覆寫並提示重新載入 |
| 安全機制 | 首次寫回前在原檔同層 `備份/` 留一份；日期輸入可用 `115/09/10`、`2026-09-10`、`9/10`；金額可含 `$`、`,` |

規則沿用帳務系統鐵則：沿用原檔名、不覆寫公式、不刪列（只清空）、大檔 `load_workbook` 需 2–5 分鐘（啟動時等待即可）。

```bash
python3 tests/make_fund_plan_sample.py      # 產生同構範例 data/月資金計劃表_範例.xlsx 供試用
python3 -m unittest tests/test_fund_plan.py  # 模型／公式評估器／API 測試
```
