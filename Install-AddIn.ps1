# ============================================================
#  Install-AddIn.ps1
#  WordAutoSaveAddin – Offline Installation Helper
#
#  Run this script from an elevated (or normal user) PowerShell
#  prompt AFTER building the project in Visual Studio.
#
#  Usage:
#    .\Install-AddIn.ps1 -BuildOutputPath "C:\...\bin\Release"
# ============================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$BuildOutputPath
)

# ── Validate path ────────────────────────────────────────────
$vstoFile = Join-Path $BuildOutputPath "WordAutoSaveAddin.vsto"
if (-not (Test-Path $vstoFile)) {
    Write-Error "Cannot find '$vstoFile'. Build the project first."
    exit 1
}

# ── Registry key ────────────────────────────────────────────
$regKey = "HKCU:\SOFTWARE\Microsoft\Office\Word\Addins\WordAutoSaveAddin"

# Manifest value uses the vstolocal suffix to bypass ClickOnce
$manifestValue = "$vstoFile|vstolocal"

Write-Host ""
Write-Host "Registering Word Auto-Save Add-in..." -ForegroundColor Cyan
Write-Host "  Build path : $BuildOutputPath"
Write-Host "  Manifest   : $manifestValue"
Write-Host ""

if (-not (Test-Path $regKey)) {
    New-Item -Path $regKey -Force | Out-Null
}

Set-ItemProperty -Path $regKey -Name "Description"   -Value "Automatically saves the active Word document every 10 seconds."
Set-ItemProperty -Path $regKey -Name "FriendlyName"  -Value "Word Auto-Save Add-in"
Set-ItemProperty -Path $regKey -Name "LoadBehavior"  -Value 3 -Type DWord
Set-ItemProperty -Path $regKey -Name "Manifest"      -Value $manifestValue

Write-Host "Registration complete." -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "  1. Launch (or restart) Microsoft Word 2019."
Write-Host "  2. Look for the 'Auto-Save' tab in the ribbon."
Write-Host "  3. If Word shows a security prompt, click 'Enable'."
Write-Host ""
