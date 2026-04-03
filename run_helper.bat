@echo off
rem Keep this file ASCII-only: Chinese Windows CMD often mis-parses UTF-8 .bat and tries to run garbage as commands.
chcp 65001 >nul
title Practice2 Setup Helper
cls

echo ========================================================
echo   Practice2 - Graphical setup / diagnostics (recommended for first run)
echo ========================================================
echo.

cd /d "%~dp0"

if exist ".venv\Scripts\activate.bat" (
    echo Found .venv, activating...
    call .venv\Scripts\activate.bat
) else if exist "venv\Scripts\activate.bat" (
    echo Found venv, activating...
    call venv\Scripts\activate.bat
) else (
    echo No venv found, using system Python...
)

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Python not found. Install Python 3.10+ and enable "Add Python to PATH".
    echo.
    pause
    exit /b 1
)

cd src
if not exist "startup_helper_gui.py" (
    echo.
    echo [ERROR] Missing src\startup_helper_gui.py
    echo Current directory: %CD%
    echo.
    pause
    exit /b 1
)

python startup_helper_gui.py
if errorlevel 1 (
    echo.
    echo [FAIL] Helper exited with an error. See messages above.
    echo If a dialog appeared, check helper_last_error.log in the project root.
    echo If tkinter is missing, reinstall Python from python.org (include Tcl/Tk).
    echo.
    pause
)
