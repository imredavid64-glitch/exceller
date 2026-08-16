#===============================================================================
# Excel Assistant - One-Click Installer (PowerShell)
#
# Compiles the add-in (.xlam) from source, installs it into Excel's AddIns
# folder, and registers it so Excel loads it automatically. No manual VBA
# steps and no administrator rights required (HKCU only).
#
# Usage: right-click > Run with PowerShell, or:
#        powershell -ExecutionPolicy Bypass -File install.ps1
#===============================================================================

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Excel Assistant - One-Click Installer" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Allow Excel macros to run and let the installer compile the add-in (HKCU only)
foreach ($ver in 16, 15, 14) {
    $key = "HKCU:\Software\Microsoft\Office\$ver.0\Excel\Security"
    if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
    New-ItemProperty -Path $key -Name "VBAWarnings" -Value 1 -PropertyType DWord -Force | Out-Null
}

$builder = Join-Path $PSScriptRoot "..\ExcellerInstaller\build_addin.vbs"
$basFile = Join-Path $PSScriptRoot "ExcelAssistant.bas"
$frmFile = Join-Path $PSScriptRoot "frmAssistant.frm"

if (-not (Test-Path $builder)) {
    Write-Host "ERROR: Could not find the installer engine:" -ForegroundColor Red
    Write-Host "  $builder" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "Building and installing the add-in..." -ForegroundColor Yellow
& cscript.exe //nologo $builder $basFile $frmFile "ExcelAssistant"
if ($LASTEXITCODE -ne 0) {
    Write-Host "" -ForegroundColor Red
    Write-Host "Installation FAILED. See the error message above." -ForegroundColor Red
    Write-Host "Tip: close Microsoft Excel, then run this installer again." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "Launching Excel - Excel Assistant will load automatically..." -ForegroundColor Green
Start-Process excel

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  INSTALLATION COMPLETE!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Excel Assistant is installed and will load every time Excel starts." -ForegroundColor White
Write-Host 'Look for the "Excel Assistant" toolbar at the top of Excel.' -ForegroundColor White
Write-Host ""
Read-Host "Press Enter to exit"
