@echo off
setlocal
cd /d "%~dp0"
title Hover Folder Size

if not exist ".venv\Scripts\python.exe" (
    echo [1/2] Setting up Python environment...
    where py >nul 2>nul
    if not errorlevel 1 (
        py -3 -m venv .venv
    ) else (
        python -m venv .venv
    )
    if errorlevel 1 goto :python_error
)

if not exist ".venv\.dependencies_ready" (
    echo [2/2] Installing required packages ^(first run only^)...
    ".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r requirements.txt
    if errorlevel 1 goto :install_error
    echo installed>".venv\.dependencies_ready"
)

echo.
echo Starting hover size tool. Keep this window open while you use it.
echo.
".venv\Scripts\python.exe" hover_size.py
if errorlevel 1 goto :run_error
exit /b 0

:python_error
echo.
echo Python 3 was not found. Install Python 3.10+ from python.org,
echo and tick the option to add Python to PATH.
pause
exit /b 1
:install_error
echo.
echo Dependency installation failed. Check Internet access and try again.
pause
exit /b 1
:run_error
echo.
echo The hover-size tool encountered an error.
pause
exit /b 1
