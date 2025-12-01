@echo off
echo ========================================
echo   Automation Hub - Desktop App
echo ========================================
echo.

if not exist "venv" (
    echo [1/3] Creating virtual environment...
    python -m venv venv
    if errorlevel 1 (
        echo ERROR: Python not found. Install Python 3.8+ first.
        pause
        exit /b 1
    )
)

echo [2/3] Installing dependencies...
call venv\Scripts\activate
pip install -q -r requirements.txt

echo [3/3] Starting Automation Hub...
echo.
python main.py
pause
