# 發票自動列印腳本 — 執行說明

## 環境需求
- Python 3.9 以上（Windows / Mac 均可）
- 可連線至 `line.wytc.com.tw` 的電腦（公司內網或有存取權限的網路）

---

## 步驟一：安裝依賴套件

開啟「命令提示字元」或「終端機」，執行：

```
pip install playwright openpyxl
python -m playwright install chromium
```

---

## 步驟二：準備檔案

將以下兩個檔案放在**同一個資料夾**：

| 檔案 | 說明 |
|------|------|
| `print_invoices.py` | 自動化腳本 |
| `deed46c8-______20260603.xlsx` | 您提供的 Excel 發票清單 |

---

## 步驟三：執行腳本

```
python print_invoices.py
```

執行後會：
1. 自動開啟 Chromium 瀏覽器（可見視窗）
2. 登入帳號 `01` / 密碼 `246`
3. 進入「發票作業」
4. 逐筆搜尋 Excel 清單中的 19 張發票，各存成 A4 PDF

PDF 檔案存放在同目錄的 **`invoices_pdf/`** 資料夾中。

---

## 常見問題

### Q：腳本停在登入頁面不動？
執行前先將腳本頂端的 `HEADLESS = False`（已預設），這樣可以看到瀏覽器畫面。
若欄位名稱不符，請查看 `debug_login.png` 截圖，告知我頁面上的欄位名稱，我再調整腳本。

### Q：找不到「發票作業」選單？
腳本會截圖 `debug_invoice_menu.png`，也會暫停並提示您手動點選後按 Enter 繼續。

### Q：某幾筆發票找不到？
失敗的筆數會存截圖（`invoices_pdf/001_error.png` 等），可對照截圖確認原因。

### Q：如何背景執行（不顯示視窗）？
確認流程正確後，將腳本第 17 行改為：
```python
HEADLESS = True
```

---

## 輸出檔案格式

```
invoices_pdf/
  001_AW5256253.pdf      ← 堯鑫鋼鐵有限公司
  002_TD2373073.pdf
  003_DG1991485.pdf
  ...
  019_TD6283135.pdf
```
