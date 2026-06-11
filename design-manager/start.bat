@echo off
chcp 65001 >nul 2>&1
echo ========================================
echo   Design Project Management System
echo ========================================
echo.

REM Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python not found. Please install Python 3.8+
    pause
    exit /b 1
)

for /f "delims=" %%a in ('python --version 2^>^&1') do set PYVER=%%a
echo Python Version: %PYVER%
echo.

REM Create virtual environment if not exists or incomplete
if not exist "venv\Scripts\activate.bat" (
    if exist "venv" (
        echo Removing incomplete virtual environment...
        rmdir /s /q venv
    )
    echo Creating virtual environment...
    python -m venv venv
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo Warning: Failed to activate venv, using system Python
) else (
    echo Virtual environment activated successfully
)

REM Install dependencies
echo Installing dependencies...
pip install -r requirements.txt -q
if %errorlevel% neq 0 (
    echo Warning: Some dependencies may not be installed
)

echo.
echo ========================================
echo   Starting server...
echo ========================================
echo.
echo   Local: http://localhost:5000
echo.
echo   Default account: admin / admin123
echo   Press Ctrl+C to stop
echo.
echo ========================================
echo.

REM Start server
python app.py
pause