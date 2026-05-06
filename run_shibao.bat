@echo off
chcp 65001 >nul 2>&1
cd /d %~dp0
set "PYTHONPATH=%~dp0"
"%~dp0venv\Scripts\python.exe" "%~dp0main.py" shibao
