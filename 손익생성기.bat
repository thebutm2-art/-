@echo off
chcp 65001 >nul
cd /d "%~dp0"
start "" pythonw "손익생성기.pyw"
