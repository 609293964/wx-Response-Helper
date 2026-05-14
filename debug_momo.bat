@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Missing .venv\Scripts\python.exe
    echo Run dependency installation first.
    pause
    exit /b 1
)

".venv\Scripts\python.exe" "wechat_gui_momo.py"
echo.
echo Application exited with code %ERRORLEVEL%.
pause
