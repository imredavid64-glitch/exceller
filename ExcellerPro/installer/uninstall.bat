@echo off
title Exceller Pro Uninstaller
color 0C

echo.
echo  ====================================================
echo           EXCELLER PRO - UNINSTALLER
echo  ====================================================
echo.

set ADDIN_DIR=%APPDATA%\Microsoft\AddIns

echo  This will remove Exceller Pro from your computer.
echo.
set /p confirm="Are you sure? (Y/N): "
if /i not "%confirm%"=="Y" (
    echo  Cancelled.
    pause
    exit /b 0
)
echo.

REM Close Excel if open
echo  [*] Checking for Excel...
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>nul | find /i "EXCEL.EXE" >nul
if %errorlevel% equ 0 (
    echo  [!] Excel is running
    echo      Please close Excel first
    echo.
    set /p close="Close Excel now? (Y/N): "
    if /i "%close%"=="Y" (
        taskkill /F /IM EXCEL.EXE >nul 2>&1
        timeout /t 2 >nul
        echo  [OK] Excel closed
    ) else (
        echo  Please close Excel manually and run this again
        pause
        exit /b 1
    )
)
echo.

REM Remove add-in file
echo  [*] Removing add-in file...
if exist "%ADDIN_DIR%\ExcellerPro.xlam" (
    del /f "%ADDIN_DIR%\ExcellerPro.xlam" >nul 2>&1
    echo  [OK] Add-in file removed
) else (
    echo  [~] Add-in file not found (may already be removed)
)
echo.

REM Remove registry entries
echo  [*] Cleaning up registry...
reg delete "HKCU\Software\Microsoft\Office\16.0\Excel\Options" /v OPEN /f >nul 2>&1
reg delete "HKCU\Software\ExcellerPro" /f >nul 2>&1
echo  [OK] Registry cleaned
echo.

REM Remove shortcuts
echo  [*] Removing shortcuts...
if exist "%USERPROFILE%\Desktop\Exceller Pro.lnk" del /f "%USERPROFILE%\Desktop\Exceller Pro.lnk" >nul 2>&1
echo  [OK] Shortcuts removed
echo.

echo  ====================================================
echo           UNINSTALL COMPLETE!
echo  ====================================================
echo.
echo  Exceller Pro has been removed from your computer.
echo  Your Excel files are unchanged.
echo.
echo  ====================================================
echo.
pause
