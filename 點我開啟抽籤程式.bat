@echo off
chcp 65001 >nul
if exist "%~dp0index.html" (
    start "" "%~dp0index.html"
) else (
    start "" "%~dp0抽籤程式.html"
)
