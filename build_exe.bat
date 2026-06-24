@echo off
REM ============================================================
REM  Build SegmentPro.exe for distribution
REM  Run this on a Windows machine with Python installed.
REM  The resulting .exe in dist\ can be copied to any Windows
REM  machine — no Python install required on end-user machines.
REM ============================================================

cd /d "%~dp0"

echo.
echo Installing build dependencies...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller

echo.
echo Building SegmentPro.exe...
pyinstaller --onefile --windowed ^
    --name SegmentPro ^
    --clean ^
    SegmentPro.py

echo.
if exist dist\SegmentPro.exe (
    echo Build complete. The executable is at:
    echo    %CD%\dist\SegmentPro.exe
    echo.
    echo You can copy SegmentPro.exe to any Windows machine and run it.
) else (
    echo Build failed. Check the output above for errors.
)
echo.
pause
