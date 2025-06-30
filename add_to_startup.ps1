$ScriptPath = $MyInvocation.MyCommand.Path
$AppDirectory = Split-Path $ScriptPath -Parent
$MainScriptPath = Join-Path $AppDirectory "main.py"
$StartupFolder = [Environment]::GetFolderPath("Startup")
$ShortcutPath = Join-Path $StartupFolder "PPT Redirector.lnk"

# Create a direct shortcut to pythonw with main.py as argument to completely hide window
$PythonwPath = "pythonw.exe"  # Using pythonw in PATH

# Create a shortcut in the Windows Startup folder
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $PythonwPath
$Shortcut.Arguments = "`"$MainScriptPath`""
$Shortcut.Description = "Start PPT Redirector"
$Shortcut.WorkingDirectory = $AppDirectory
$Shortcut.WindowStyle = 0  # Hidden (0), Normal (1), Minimized (7)
$Shortcut.Save()

if (Test-Path $ShortcutPath) {
    Write-Host "Startup shortcut created successfully at: $ShortcutPath"
    Write-Host "PPT Redirector will now start automatically when you log in to Windows."
} else {
    Write-Host "Failed to create the startup shortcut."
}
