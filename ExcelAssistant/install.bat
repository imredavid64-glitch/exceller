@echo off
title Excel Assistant - One-Click Installer
setlocal
echo ========================================
echo   Excel Assistant - One-Click Installer
echo ========================================
echo.

REM Allow Excel macros to run without prompts and let the installer
REM compile the add-in automatically. (HKCU only - no admin needed.)
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v VBAWarnings /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\15.0\Excel\Security" /v VBAWarnings /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\14.0\Excel\Security" /v VBAWarnings /t REG_DWORD /d 1 /f >nul 2>&1

echo Building and installing the add-in...
cscript //nologo "%~dp0..\ExcellerInstaller\build_addin.vbs" "%~dp0ExcelAssistant.bas" "%~dp0frmAssistant.frm" "ExcelAssistant"
if %errorlevel% neq 0 (
    echo.
    echo Installation FAILED. See the error message above.
    echo Tip: close Microsoft Excel, then run this installer again.
    pause
    exit /b 1
)

echo.
echo Launching Excel - Excel Assistant will load automatically...
start "" excel

echo.
echo ========================================
echo   INSTALLATION COMPLETE!
echo ========================================
echo.
echo Excel Assistant is installed and will load every time Excel starts.
echo Look for the "Excel Assistant" toolbar at the top of Excel.
echo.
pause
endlocal
