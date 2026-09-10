@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 月資金計劃表 線上表單
echo 啟動月資金計劃表線上表單...（大檔載入需 2-5 分鐘，請勿關閉此視窗）
set PY=python
where py >nul 2>&1 && set PY=py -3
if "%~1"=="" (
  %PY% src\fund_plan_server.py --open
) else (
  %PY% src\fund_plan_server.py "%~1" --open
)
pause
