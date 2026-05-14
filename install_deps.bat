@echo off
setlocal
cd /d "%~dp0"

echo EasyChat Momo dependency installer
echo.

if exist ".venv\Scripts\python.exe" (
    echo Using existing .venv
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    goto done
)

where py >nul 2>nul
if %ERRORLEVEL%==0 (
    echo Creating .venv with py launcher
    py -3 -m venv .venv
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    goto done
)

where python >nul 2>nul
if %ERRORLEVEL%==0 (
    echo Creating .venv with python
    python -m venv .venv
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    goto done
)

echo Python was not found.
echo Install Python 3.10+ first, then run this script again.
pause
exit /b 1

:done
echo.
echo Dependencies are ready.
pause
