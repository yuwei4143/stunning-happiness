# 打包成單一執行檔（給不熟指令的同事雙擊即用）

工具提供拖放式圖形介面 `src/ap_gui.py`，可用 PyInstaller 打包成單一執行檔。

## 重要：不能跨平台編譯

PyInstaller **只能產出「執行打包指令的那個作業系統」的執行檔**：

| 你在哪個系統打包 | 產出 |
|------------------|------|
| Windows | `應付帳款預填.exe` |
| macOS | `應付帳款預填`（可再包成 .app） |
| Linux | `應付帳款預填`（ELF binary） |

要給 Windows 同事的 `.exe`，就必須在一台 Windows 上執行下面的步驟。

## 打包步驟（在目標系統上）

```bash
# 1. 安裝 Python 3.9+（Windows 請到 python.org，安裝時勾選 Add to PATH）

# 2. 安裝打包所需套件
pip install pyinstaller tkinterdnd2 openpyxl xlrd

# 3. 在專案根目錄執行打包腳本
python packaging/build_app.py
```

產出在 `dist/` 目錄。

## 交付給同事

把 `dist/` 內的執行檔，連同 **`customer_master.xls`（客戶主檔）** 放在同一個資料夾，
一起交給同事。程式會自動偵測同資料夾的主檔。

```
給同事的資料夾/
  應付帳款預填.exe        ← 雙擊執行
  customer_master.xls     ← 客戶主檔（放旁邊即自動載入）
```

## 同事的操作（三步）

1. 雙擊執行檔開啟視窗。
2. 把「支票辨識紀錄 .xlsx」拖進視窗（或按「選擇檔案」），視需要選用途／送票資訊。
3. 按「開始產出」，完成後可直接開啟輸出資料夾。輸出檔為 `<原檔名>_待回填.xlsx`。

需人工覆核的列會在最後兩欄標明，確認後貼入 11505 應付帳款明細表。

## 沒有 Windows 電腦？用 GitHub Actions 自動建置

專案內含 `.github/workflows/build-windows.yml`，會在**雲端 Windows 機器**自動打包，
不必自己準備 Windows：

1. **推送分支即自動建置**：每次推送到 `claude/**` 分支或手動觸發後，到 GitHub 專案的
   **Actions** 分頁 → 點該次執行 → 下方 **Artifacts** 下載 `應付帳款預填-windows`（內含 .exe）。
2. **手動觸發**：Actions 分頁 → 左側「打包 Windows 執行檔」→ **Run workflow**。
3. **正式發佈**：打一個 `v` 開頭的 tag（例如 `git tag v1.0 && git push origin v1.0`），
   會自動把 .exe 附到對應的 **Release** 供長期下載。

下載後同樣把 .exe 與 `customer_master.xls` 放同一資料夾交給同事即可。

## 疑難排解

- **拖放沒反應**：未安裝 `tkinterdnd2` 時拖放會停用，改按「選擇檔案」即可，功能不受影響。
- **Windows SmartScreen 攔截**：未簽章的自製 exe 首次執行會警告，點「其他資訊 → 仍要執行」。
- **客戶簡稱空白**：代表未載入到主檔，確認 `customer_master.xls` 在執行檔旁，或用視窗內「選擇檔案」手動指定。
