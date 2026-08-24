@echo off
chcp 65001 >nul
cd /d "%~dp0"
set INVENTORY_PORT=4174
if exist "%~dp0.venv\Scripts\pythonw.exe" (
  start "" /b "%~dp0.venv\Scripts\pythonw.exe" "%~dp0app.py"
) else (
  py -3.12 "%~dp0app.py"
)
