@echo off
title IT Asset Tracker - Installer
color 0A
echo.
echo ============================================================
echo   IT Asset Tracker - Installation
echo ============================================================
echo.

REM Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please install Python 3.8+ from https://python.org
    echo Make sure to check "Add Python to PATH" during install.
    pause
    exit /b 1
)

echo [OK] Python found
for /f "tokens=*" %%i in ('python --version') do echo     %%i
echo.

REM Install dependencies
echo [STEP 1] Installing required packages...
python -m pip install --upgrade pip --quiet
python -m pip install flask openpyxl --quiet

if %errorlevel% neq 0 (
    echo [ERROR] Failed to install packages. Check internet connection.
    pause
    exit /b 1
)

echo [OK] Flask installed
echo [OK] openpyxl installed
echo.

REM Verify app.py exists
if not exist "%~dp0app.py" (
    echo [ERROR] app.py not found in %~dp0
    pause
    exit /b 1
)

echo [OK] Application files verified
echo.
echo ============================================================
echo   Installation Complete!
echo ============================================================
echo.
echo   To start the application, double-click:
echo   launch.bat
echo.
echo   OR run: python app.py
echo.
echo   Default Login:
echo     Admin:  admin / admin123
echo     Users:  [employee username] / pass123
echo.
pause
