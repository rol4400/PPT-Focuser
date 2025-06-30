$StartupFolder = [Environment]::GetFolderPath("Startup")
$ShortcutPath = Join-Path $StartupFolder "PPT Redirector.lnk"

if (Test-Path $ShortcutPath) {
    Remove-Item $ShortcutPath -Force
    Write-Host "PPT Redirector has been removed from startup."
} else {
    Write-Host "PPT Redirector is not set to run at startup."
}
