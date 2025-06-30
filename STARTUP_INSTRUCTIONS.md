# PPT Redirector Startup Configuration

This document explains how to set up PPT Redirector to start automatically when your computer boots up.

## Files Included

- `start_ppt_redirector.bat` - A batch file that launches the PPT Redirector application
- `add_to_startup.ps1` - A PowerShell script that adds PPT Redirector to Windows startup
- `remove_from_startup.ps1` - A PowerShell script that removes PPT Redirector from Windows startup

## Setting Up Automatic Startup

### Method 1: Using the PowerShell Script (Recommended)

1. Right-click on `add_to_startup.ps1`
2. Select "Run with PowerShell"
3. If you see a security warning, you may need to type "Y" to confirm
4. The script will create a shortcut in your Windows Startup folder

### Method 2: Manual Setup

1. Press `Win + R` on your keyboard
2. Type `shell:startup` and press Enter
3. This opens your Windows Startup folder
4. Create a shortcut to `start_ppt_redirector.bat` in this folder

### To Remove from Startup

1. Right-click on `remove_from_startup.ps1`
2. Select "Run with PowerShell"
3. The script will remove PPT Redirector from your startup programs

## Troubleshooting

- Make sure the application works correctly by first running `start_ppt_redirector.bat` directly
- If the application doesn't start on boot, check that the paths in the shortcuts are correct
- For more detailed logging, edit the batch file and remove the `@echo off` line
