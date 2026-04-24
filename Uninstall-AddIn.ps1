# ============================================================
#  Uninstall-AddIn.ps1
#  WordAutoSaveAddin – Offline Uninstall Helper
# ============================================================

$regKey = "HKCU:\SOFTWARE\Microsoft\Office\Word\Addins\WordAutoSaveAddin"

if (Test-Path $regKey) {
    Remove-Item -Path $regKey -Recurse -Force
    Write-Host "Add-in unregistered successfully." -ForegroundColor Green
} else {
    Write-Host "Add-in registry key not found – nothing to remove." -ForegroundColor Yellow
}

Write-Host "Restart Word for the change to take effect."
