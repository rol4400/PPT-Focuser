# PPT Redirector

A Windows application that redirects arrow keys and page up/down keys to a selected target window (like PowerPoint in slideshow mode) while you work in other applications.

## Features

- System tray application that runs in the background
- Visual window selector with thumbnails/previews
- Redirects Left/Right arrow keys and Page Up/Down keys to target window
- No focus stealing - your current window stays active
- Works with PowerPoint presentations, PDF viewers, and other applications

## Installation

1. Install Python 3.7 or higher
2. Install required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```
   python main.py
   ```
   Or use the provided batch file:
   ```
   start_ppt_redirector.bat
   ```

2. The application will start in the system tray (look for the icon in the bottom-right corner)

3. Click or right-click the tray icon and select "Select Window" to choose your target window

4. Browse through the list of windows (with thumbnails) and select the one you want to control

5. Click "Set as Target" to confirm your selection

## Setting Up Automatic Startup

To make PPT Redirector start automatically when your computer boots:

1. Right-click on `add_to_startup.ps1` and select "Run with PowerShell"
2. Confirm any security prompts if they appear

See `STARTUP_INSTRUCTIONS.md` for detailed instructions and manual alternatives.

6. Now when you press Left/Right arrow keys or Page Up/Down, they will be sent to your target window instead of the currently focused window

## Controls

- **Left/Right Arrow Keys**: Navigate slides backward/forward
- **Page Up/Page Down**: Navigate slides backward/forward

## System Tray Menu

- **Select Window**: Choose a new target window
- **Show Current Target**: Display which window is currently targeted
- **Quit**: Exit the application

## Requirements

- Windows OS
- Python 3.7+
- PyQt5
- pywin32
- keyboard
- Pillow (for window thumbnails)

## Troubleshooting

If you don't see window thumbnails, make sure Pillow is installed:
```
pip install Pillow
```

If keys aren't being sent to the target window, try:
1. Running as administrator
2. Ensuring the target application is responsive
3. Selecting a different window and trying again
