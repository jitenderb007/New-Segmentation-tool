@echo off
REM ============================================================
REM  SegmentPro — Windows launcher
REM  Double-click this file to run the app from source.
REM ============================================================

cd /d "%~dp0"

REM Check for Python
where python >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERROR: Python is not installed or not on PATH.
    echo Download and install Python 3.9+ from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

REM Check for required packages; install if missing
python -c "import pandas, numpy, sklearn, openpyxl, pptx" >nul 2>&1
if errorlevel 1 (
    echo.
    echo Installing required Python packages...
    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo ERROR: Could not install dependencies.
        pause
        exit /b 1
    )
)

REM Launch the app
python SegmentPro.py
