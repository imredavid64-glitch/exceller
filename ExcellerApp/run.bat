@echo off
echo Starting Exceller...

REM Check if venv exists
if not exist "venv" (
    echo Virtual environment not found. Please run install.bat first.
    pause
    exit /b 1
)

REM Activate virtual environment
call venv\Scripts\activate

REM Run the application
python main.py

pause
