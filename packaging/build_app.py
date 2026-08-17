"""用 PyInstaller 把 GUI 打包成單一執行檔。

在目標作業系統上執行（Windows 產出 .exe、macOS 產出 .app/binary、Linux 產出 binary）。
PyInstaller 無法跨平台編譯——要 Windows 執行檔就在 Windows 上跑這支腳本。

用法：
    pip install pyinstaller tkinterdnd2 openpyxl xlrd
    python packaging/build_app.py

產出在 dist/ 目錄，將該執行檔連同 customer_master.xls 一起交給同事即可雙擊使用。
"""
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ENTRY = os.path.join(ROOT, "src", "ap_gui.py")
NAME = "應付帳款預填"


def main():
    args = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--windowed",
        "--name", NAME,
        "--paths", os.path.join(ROOT, "src"),
        # 確保同層模組被收錄
        "--hidden-import", "normalize",
        "--hidden-import", "customer_master",
        "--hidden-import", "payable_filler",
        # 拖放套件的原生元件（未安裝則忽略）
        "--collect-all", "tkinterdnd2",
        "--noconfirm",
        "--distpath", os.path.join(ROOT, "dist"),
        "--workpath", os.path.join(ROOT, "build"),
        "--specpath", os.path.join(ROOT, "build"),
        ENTRY,
    ]
    print("執行：", " ".join(args))
    rc = subprocess.call(args)
    if rc == 0:
        print(f"\n完成 → {os.path.join(ROOT, 'dist', NAME)}")
        print("把該執行檔與 customer_master.xls 放同一資料夾，雙擊即用。")
    sys.exit(rc)


if __name__ == "__main__":
    main()
