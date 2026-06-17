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
    for %%V in (3.13 3.12 3.11 3.10) do (
        py -%%V -c "import sys" >nul 2>nul
        if not errorlevel 1 (
            echo Creating .venv with Python %%V
            py -%%V -m venv .venv
            goto install
        )
    )
)

where python >nul 2>nul
if %ERRORLEVEL%==0 (
    python -c "import sys; raise SystemExit(0 if (3, 10) <= sys.version_info[:2] <= (3, 13) else 1)" >nul 2>nul
    if not errorlevel 1 (
        echo Creating .venv with python
        python -m venv .venv
        goto install
    )
)

echo A compatible Python was not found.
echo Install Python 3.10 through 3.13, then run this script again.
pause
exit /b 1

:install
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
    echo Dependency installation failed.
    pause
    exit /b 1
)

:done
echo.
echo Dependencies are ready.
pause
