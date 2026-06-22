@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/tpougy/finance-fmt-tools/main/Install-RBRFinanceTools.ps1 | iex"
pause