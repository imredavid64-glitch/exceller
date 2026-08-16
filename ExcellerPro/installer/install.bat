@echo off
title Exceller Pro - One-Click Installer
setlocal
echo ========================================
echo   EXCELLER PRO - ONE-CLICK INSTALLER
echo   AI-Powered Excel Assistant
echo ========================================
echo.

REM Allow Excel macros to run and let the installer compile the add-in.
REM (HKCU only - no administrator rights needed.)
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v VBAWarnings /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\15.0\Excel\Security" /v VBAWarnings /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\14.0\Excel\Security" /v VBAWarnings /t REG_DWORD /d 1 /f >nul 2>&1

echo  [*] Building and installing Exceller Pro...
cscript //nologo "%~dp0..\..\ExcellerInstaller\build_addin.vbs" "%~dp0..\ExcellerPro.bas" "" "ExcellerPro"
if %errorlevel% neq 0 (
    echo.
    echo  [!] Installation FAILED. See the error message above.
    echo      Tip: close Microsoft Excel, then run this installer again.
    pause
    exit /b 1
)

echo.
echo  [*] Launching Excel - Exceller Pro will load automatically...
start "" excel

echo.
echo  ====================================================
echo           INSTALLATION COMPLETE!
echo  ====================================================
echo.
echo  Exceller Pro is installed and will load every time Excel starts.
echo  Use the "Exceller Pro" toolbar at the top of Excel, or press Alt+F8.
echo.
echo  GET API KEY:
echo  https://makersuite.google.com/app/apikey
echo.
echo  ====================================================
pause
endlocal
