@echo off
title IT Asset Tracker
color 0B

REM Change to script directory
cd /d "%~dp0"

REM Quick check for Flask
python -c "import flask" >nul 2>&1
if %errorlevel% neq 0 (
    echo Flask not found. Running installer first...
    call install.bat
)

echo.
echo ============================================================
echo   IT Asset Tracker is starting...
echo ============================================================
echo.
echo   URL:   http://localhost:5000
echo   Admin: admin / admin123
echo   Users: [employee username] / pass123
echo.
echo   Close this window to stop the server.
echo ============================================================
echo.

REM Open browser after 2 seconds
start "" /b cmd /c "timeout /t 2 /nobreak >nul && start http://localhost:5000"

REM Start Flask app
python app.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Application failed to start.
    echo Make sure port 5000 is not in use.
    pause
)
