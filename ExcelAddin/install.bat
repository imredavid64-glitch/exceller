@echo off
title Exceller - One-Click Installer
setlocal
echo ========================================
echo   Exceller - Excel Add-in Installer
echo ========================================
echo.

REM Allow Excel macros to run without prompts and let the installer
REM compile the add-in automatically. (HKCU only - no admin needed.)
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v VBAWarnings /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\15.0\Excel\Security" /v VBAWarnings /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\14.0\Excel\Security" /v VBAWarnings /t REG_DWORD /d 1 /f >nul 2>&1

echo Building and installing the add-in...
cscript //nologo "%~dp0..\ExcellerInstaller\build_addin.vbs" "%~dp0ExcelAssistant.bas" "" "ExcelAssistant"
if %errorlevel% neq 0 (
    echo.
    echo Installation FAILED. See the error message above.
    echo Tip: close Microsoft Excel, then run this installer again.
    pause
    exit /b 1
)

echo.
echo Launching Excel - Exceller will load automatically...
start "" excel

echo.
echo ========================================
echo   INSTALLATION COMPLETE!
echo ========================================
echo.
echo Exceller is installed and will load every time Excel starts.
echo In Excel: press Alt+F8 to see all Exceller commands.
echo.
pause
endlocal
