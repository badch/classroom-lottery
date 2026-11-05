@echo off
chcp 65001 >nul
title 班級抽籤系統 - Python 版

where python >nul 2>nul
if %errorlevel% equ 0 (
    python "%~dp0抽籤程式.py"
    if %errorlevel% neq 0 pause
    exit /b
)

where py >nul 2>nul
if %errorlevel% equ 0 (
    py "%~dp0抽籤程式.py"
    if %errorlevel% neq 0 pause
    exit /b
)

echo ========================================================
echo [提示] 尚未偵測到 Python 環境！
echo.
echo 建議直接使用「點我開啟抽籤程式.bat」開啟網頁版（功能相同，免安裝）。
echo.
echo 若您想執行 Python 版：
echo 1. 請至 https://www.python.org/downloads/ 下載安裝
echo 2. 安裝時務必勾選「Add python.exe to PATH」
echo ========================================================
echo.
pause
