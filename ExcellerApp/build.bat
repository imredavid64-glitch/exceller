@echo off
echo ========================================
echo Building Exceller - Standalone App
echo ========================================
echo.

REM Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python not found! Install from python.org
    pause
    exit /b 1
)

REM Install dependencies
echo Installing dependencies...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo Building Exceller...
echo This may take 2-3 minutes...
echo.

pyinstaller --onefile --windowed --name "Exceller" main.py

echo.
echo ========================================
echo Build Complete!
echo ========================================
echo.
echo Find Exceller.exe in the "dist" folder
echo.
pause
