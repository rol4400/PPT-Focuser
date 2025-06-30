$ScriptPath = $MyInvocation.MyCommand.Path
$AppDirectory = Split-Path $ScriptPath -Parent
$BatchFilePath = Join-Path $AppDirectory "start_ppt_redirector.bat"
$StartupFolder = [Environment]::GetFolderPath("Startup")
$ShortcutPath = Join-Path $StartupFolder "PPT Redirector.lnk"

# Create a shortcut to the batch file in the Windows Startup folder
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $BatchFilePath
$Shortcut.Description = "Start PPT Redirector"
$Shortcut.WorkingDirectory = $AppDirectory
$Shortcut.WindowStyle = 7  # Minimized
$Shortcut.Save()

if (Test-Path $ShortcutPath) {
    Write-Host "Startup shortcut created successfully at: $ShortcutPath"
    Write-Host "PPT Redirector will now start automatically when you log in to Windows."
} else {
    Write-Host "Failed to create the startup shortcut."
}
