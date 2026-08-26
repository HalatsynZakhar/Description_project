@echo off
cd /d "%~dp0"
".venv\Scripts\python.exe" "check_keys.py"
echo.
pause
