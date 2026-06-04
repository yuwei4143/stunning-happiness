"""
自動登入 line.wytc.com.tw 並將 Excel 中每筆發票另存為 A4 PDF
使用方式：
  1. 安裝依賴：pip install playwright openpyxl
  2. 安裝瀏覽器：python -m playwright install chromium
  3. 執行：python print_invoices.py
"""

import asyncio
import sys
import os
from pathlib import Path
from datetime import datetime

try:
    import openpyxl
except ImportError:
    print("請先執行：pip install openpyxl")
    sys.exit(1)

try:
    from playwright.async_api import async_playwright, TimeoutError as PWTimeout
except ImportError:
    print("請先執行：pip install playwright && python -m playwright install chromium")
    sys.exit(1)

# ── 設定區 ────────────────────────────────────────────────────────────────────
EXCEL_PATH   = "deed46c8-______20260603.xlsx"   # Excel 檔案路徑（可修改）
OUTPUT_DIR   = "invoices_pdf"                   # PDF 輸出資料夾
LOGIN_URL    = "https://line.wytc.com.tw/st2/login.aspx"
USERNAME     = "01"
PASSWORD     = "246"
HEADLESS     = False   # True=背景執行；False=顯示瀏覽器視窗（建議先設 False 確認流程）
# ─────────────────────────────────────────────────────────────────────────────


def load_excel(path: str) -> list[dict]:
    """從 Excel 讀取支票/發票清單，回傳 list of dict"""
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        records.append(dict(zip(headers, row)))
    print(f"[Excel] 讀取到 {len(records)} 筆資料")
    return records


async def login(page, url: str, username: str, password: str) -> bool:
    """登入網站，成功回傳 True"""
    print(f"[登入] 前往 {url}")
    await page.goto(url, wait_until="networkidle", timeout=30000)

    # 嘗試常見的帳號/密碼欄位選擇器（ASP.NET 網站通常有 ViewState）
    selectors_user = [
        "input[name*='user' i]", "input[name*='User' i]",
        "input[name*='account' i]", "input[name*='id' i]",
        "input[type='text']:first-of-type", "#txtUserID", "#TextBox1",
    ]
    selectors_pass = [
        "input[type='password']", "input[name*='pass' i]",
        "input[name*='pwd' i]", "#txtPassword", "#TextBox2",
    ]
    selectors_btn = [
        "input[type='submit']", "button[type='submit']",
        "input[value*='登入']", "input[value*='Login']",
        "#btnLogin", "#Button1",
    ]

    filled_user = False
    for sel in selectors_user:
        try:
            await page.fill(sel, username, timeout=3000)
            print(f"  帳號欄：{sel}")
            filled_user = True
            break
        except Exception:
            pass

    filled_pass = False
    for sel in selectors_pass:
        try:
            await page.fill(sel, password, timeout=3000)
            print(f"  密碼欄：{sel}")
            filled_pass = True
            break
        except Exception:
            pass

    if not (filled_user and filled_pass):
        # 拍截圖方便人工查看
        await page.screenshot(path="debug_login.png")
        print("[錯誤] 找不到帳號/密碼欄位，已截圖至 debug_login.png，請手動確認頁面結構")
        return False

    for sel in selectors_btn:
        try:
            await page.click(sel, timeout=3000)
            print(f"  按下登入按鈕：{sel}")
            break
        except Exception:
            pass

    await page.wait_for_load_state("networkidle", timeout=20000)
    await page.screenshot(path="debug_after_login.png")
    print("[登入] 登入後截圖已存至 debug_after_login.png")
    return True


async def navigate_to_invoice_menu(page) -> bool:
    """嘗試進入「發票作業」選單"""
    keywords = ["發票作業", "發票", "Invoice", "invoice"]
    for kw in keywords:
        try:
            locator = page.get_by_text(kw, exact=False)
            count = await locator.count()
            if count > 0:
                await locator.first.click(timeout=5000)
                await page.wait_for_load_state("networkidle", timeout=15000)
                print(f"[選單] 點擊「{kw}」")
                await page.screenshot(path="debug_invoice_menu.png")
                return True
        except Exception:
            pass
    await page.screenshot(path="debug_invoice_menu.png")
    print("[選單] 找不到發票作業項目，截圖已存至 debug_invoice_menu.png")
    return False


async def search_and_print_invoice(page, record: dict, output_dir: Path, index: int) -> bool:
    """
    依支票號碼或客戶名稱搜尋發票並列印成 PDF
    實際搜尋邏輯需依網站介面調整。
    """
    ticket_no  = record.get("支票號碼", "")
    customer   = record.get("客戶") or record.get("發票人", "")
    amount     = record.get("實收金額", "")
    due_date   = record.get("到期日", "")
    if isinstance(due_date, datetime):
        due_date = due_date.strftime("%Y/%m/%d")

    label = ticket_no or customer or f"row{index}"
    print(f"\n[{index}] 處理：{label}  客戶：{customer}  金額：{amount}")

    # ── 嘗試在搜尋欄輸入支票號碼 ──────────────────────────────────────────
    search_selectors = [
        "input[name*='search' i]", "input[name*='key' i]",
        "input[name*='no' i]", "input[name*='num' i]",
        "input[type='text']",
    ]
    searched = False
    for sel in search_selectors:
        try:
            elements = await page.query_selector_all(sel)
            if elements:
                await elements[0].triple_click()
                await elements[0].type(ticket_no, delay=50)
                searched = True
                print(f"  搜尋欄輸入：{ticket_no}")
                break
        except Exception:
            pass

    if searched:
        # 按下查詢按鈕
        query_btns = ["input[value*='查詢']", "input[value*='搜尋']",
                      "button:has-text('查詢')", "input[type='submit']"]
        for sel in query_btns:
            try:
                await page.click(sel, timeout=3000)
                await page.wait_for_load_state("networkidle", timeout=15000)
                break
            except Exception:
                pass

    # ── 嘗試點擊第一筆結果 ────────────────────────────────────────────────
    row_selectors = [
        f"tr:has-text('{ticket_no}')",
        f"td:has-text('{ticket_no}')",
        "table tbody tr:nth-child(1)",
    ]
    for sel in row_selectors:
        try:
            await page.click(sel, timeout=4000)
            await page.wait_for_load_state("networkidle", timeout=10000)
            print(f"  點擊結果列：{sel}")
            break
        except Exception:
            pass

    # ── 列印成 PDF（A4） ──────────────────────────────────────────────────
    safe_label = "".join(c if c.isalnum() or c in "-_" else "_" for c in label)
    pdf_path = output_dir / f"{index:03d}_{safe_label}.pdf"

    try:
        await page.pdf(
            path=str(pdf_path),
            format="A4",
            print_background=True,
            margin={"top": "10mm", "bottom": "10mm",
                    "left": "10mm", "right": "10mm"},
        )
        print(f"  PDF 已存至：{pdf_path}")
        return True
    except Exception as e:
        # 截圖備用
        png_path = output_dir / f"{index:03d}_{safe_label}.png"
        await page.screenshot(path=str(png_path), full_page=True)
        print(f"  [警告] PDF 失敗（{e}），改存截圖：{png_path}")
        return False


async def main():
    # 確認 Excel 存在
    if not Path(EXCEL_PATH).exists():
        print(f"[錯誤] 找不到 Excel 檔案：{EXCEL_PATH}")
        print("請將 Excel 檔案放在與本腳本相同目錄，或修改腳本頂端的 EXCEL_PATH")
        sys.exit(1)

    records = load_excel(EXCEL_PATH)

    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(exist_ok=True)
    print(f"[輸出] PDF 將存至資料夾：{output_dir.resolve()}")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=HEADLESS,
            args=["--no-sandbox", "--disable-setuid-sandbox"],
        )
        context = await browser.new_context(
            viewport={"width": 1280, "height": 900},
            locale="zh-TW",
        )
        page = await context.new_page()

        # 1. 登入
        if not await login(page, LOGIN_URL, USERNAME, PASSWORD):
            print("[中止] 登入失敗，請查看截圖確認頁面結構後手動調整腳本")
            await browser.close()
            return

        # 2. 進入發票作業選單
        menu_ok = await navigate_to_invoice_menu(page)
        if not menu_ok:
            print("[警告] 未自動找到發票作業選單，請確認截圖後繼續")
            input("請手動操作瀏覽器到發票作業頁面後，按 Enter 繼續...")

        # 3. 逐筆列印
        success, failed = 0, 0
        for i, record in enumerate(records, start=1):
            try:
                ok = await search_and_print_invoice(page, record, output_dir, i)
                if ok:
                    success += 1
                else:
                    failed += 1
            except Exception as e:
                print(f"  [錯誤] 第 {i} 筆例外：{e}")
                failed += 1
                await page.screenshot(path=str(output_dir / f"{i:03d}_error.png"))

        await browser.close()

    print(f"\n═══════════════════════════════════════")
    print(f"完成！成功：{success} 筆 ／ 失敗/警告：{failed} 筆")
    print(f"PDF 存放位置：{Path(OUTPUT_DIR).resolve()}")
    print(f"═══════════════════════════════════════")


if __name__ == "__main__":
    asyncio.run(main())
