@echo off
echo ========================================
echo   Automation Hub - Anaconda Version
echo ========================================
echo.

REM Check if conda is available
where conda >nul 2>nul
if errorlevel 1 (
    echo ERROR: Conda not found. Please install Anaconda or Miniconda first.
    echo Download from: https://www.anaconda.com/download
    pause
    exit /b 1
)

REM Check if environment exists
conda env list | findstr "automation_hub" >nul
if errorlevel 1 (
    echo [1/3] Creating conda environment 'automation_hub'...
    conda create -n automation_hub python=3.10 -y
    if errorlevel 1 (
        echo ERROR: Failed to create conda environment.
        pause
        exit /b 1
    )
) else (
    echo [1/3] Conda environment 'automation_hub' already exists.
)

echo [2/3] Activating environment and installing dependencies...
call conda activate automation_hub
pip install -q -r requirements.txt

echo [3/3] Starting Automation Hub...
echo.
python main.py

REM Deactivate when done
call conda deactivate
pause
