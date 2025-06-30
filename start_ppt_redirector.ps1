# Check if PPT Redirector is already running
$ProcessName = "pythonw"  # The process name that runs the script

# Check if there's any pythonw process running our main.py script
$isRunning = Get-WmiObject Win32_Process | 
             Where-Object { $_.Name -eq "$ProcessName.exe" -and $_.CommandLine -like "*main.py*" }

if ($isRunning) {
    Write-Host "PPT Redirector is already running."
} else {
    Write-Host "Starting PPT Redirector..."
    
    $ScriptPath = $MyInvocation.MyCommand.Path
    $AppDirectory = Split-Path $ScriptPath -Parent
    $MainScript = Join-Path $AppDirectory "main.py"
    
    # Start pythonw.exe directly with main.py as the argument, completely hidden
    Start-Process -FilePath "pythonw.exe" -ArgumentList "`"$MainScript`"" -WindowStyle Hidden -WorkingDirectory $AppDirectory
    
    Write-Host "PPT Redirector has been started."
}
