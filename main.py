import sys
import os
import time
import ctypes
import threading
import subprocess
from io import BytesIO
import win32gui
import win32con
import win32process
import keyboard
import winreg

# Try to import optional dependencies for thumbnails
try:
    import win32ui
    import win32api
    from PIL import Image
    from PyQt5.QtCore import QBuffer
    # PIL.ImageQt is broken in newer versions, we'll use manual conversion
    THUMBNAIL_SUPPORT = True
except ImportError:
    print("Warning: PIL (Pillow) and/or pywin32 extensions not found. Window thumbnails will be limited to icons only.")
    THUMBNAIL_SUPPORT = False

def pil_to_qpixmap(pil_image):
    """Convert PIL Image to QPixmap - workaround for broken PIL.ImageQt"""
    try:
        # Convert PIL image to bytes
        from io import BytesIO
        from PyQt5.QtGui import QImage
        
        # Convert to RGB if not already
        if pil_image.mode != 'RGB':
            pil_image = pil_image.convert('RGB')
        
        # Use in-memory PNG conversion to ensure proper orientation
        buffer = BytesIO()
        pil_image.save(buffer, format='PNG')
        buffer.seek(0)
        
        # Load from PNG data directly
        qimage = QImage()
        qimage.loadFromData(buffer.getvalue())
        
        # Convert to QPixmap
        pixmap = QPixmap.fromImage(qimage)
        return pixmap
        
    except Exception as e:
        print(f"PIL to QPixmap conversion failed: {e}")
        return None

from PyQt5.QtWidgets import (
    QApplication, QWidget, QListWidget, QVBoxLayout, QPushButton,
    QSystemTrayIcon, QMenu, QAction, QMessageBox, QListWidgetItem,
    QHBoxLayout, QLabel
)
from PyQt5.QtGui import QIcon, QPixmap, QPainter
from PyQt5.QtCore import QTimer, Qt, QSize

def convert_hicon_to_qpixmap(hicon, size=(64, 64)):
    """Convert Windows HICON to QPixmap using win32gui bitmap operations"""
    try:
        if not hicon:
            return None
            
        # Get device context
        hdc = win32gui.GetDC(0)
        hdc_mem = win32gui.CreateCompatibleDC(hdc)
        
        # Create a bitmap
        hbmp = win32gui.CreateCompatibleBitmap(hdc, size[0], size[1])
        win32gui.SelectObject(hdc_mem, hbmp)
        
        # Fill with white background
        win32gui.PatBlt(hdc_mem, 0, 0, size[0], size[1], win32con.WHITENESS)
        
        # Draw the icon onto the bitmap
        win32gui.DrawIconEx(hdc_mem, 0, 0, hicon, size[0], size[1], 0, 0, 0x0003)  # DI_NORMAL | DI_COMPAT
        
        if THUMBNAIL_SUPPORT:
            try:
                # Get bitmap bits using win32ui
                import win32ui
                
                # Create device context objects
                mfc_dc = win32ui.CreateDCFromHandle(hdc_mem)
                save_bitmap = win32ui.CreateBitmapFromHandle(hbmp)
                
                # Get bitmap info and bits
                bmp_info = save_bitmap.GetInfo()
                bmp_bits = save_bitmap.GetBitmapBits(True)
                
                # Use the actual bitmap stride from Windows
                stride = bmp_info['bmWidthBytes']
                width = bmp_info['bmWidth']
                height = bmp_info['bmHeight']
                
                # Convert icon bitmap directly to in-memory BMP file to avoid orientation issues
                from io import BytesIO
                import struct
                
                # Create BMP file in memory
                bmp_file = BytesIO()
                
                # BMP file header (14 bytes)
                bmp_file.write(b'BM')  # Signature
                file_size = 14 + 40 + len(bmp_bits)  # 14=header, 40=DIB header size
                bmp_file.write(struct.pack('<I', file_size))  # FileSize
                bmp_file.write(struct.pack('<I', 0))  # Reserved
                bmp_file.write(struct.pack('<I', 14 + 40))  # DataOffset
                
                # DIB header (40 bytes - BITMAPINFOHEADER)
                bmp_file.write(struct.pack('<I', 40))  # HeaderSize
                bmp_file.write(struct.pack('<i', width))  # Width
                bmp_file.write(struct.pack('<i', -height))  # Height (negative = top-down)
                bmp_file.write(struct.pack('<H', 1))  # Planes
                bmp_file.write(struct.pack('<H', 32))  # BitCount
                bmp_file.write(struct.pack('<I', 0))  # Compression
                bmp_file.write(struct.pack('<I', len(bmp_bits)))  # ImageSize
                bmp_file.write(struct.pack('<i', 0))  # XPelsPerMeter
                bmp_file.write(struct.pack('<i', 0))  # YPelsPerMeter
                bmp_file.write(struct.pack('<I', 0))  # ClrUsed
                bmp_file.write(struct.pack('<I', 0))  # ClrImportant
                
                # Bitmap bits
                bmp_file.write(bmp_bits)
                bmp_file.seek(0)
                
                # Create image directly from memory BMP file
                img = Image.open(bmp_file)
                
                # Convert to Qt
                pixmap = pil_to_qpixmap(img)
                
                # Cleanup
                win32gui.DeleteObject(hbmp)
                win32gui.DeleteDC(hdc_mem)
                win32gui.ReleaseDC(0, hdc)
                
                return pixmap
                
            except Exception as e:
                print(f"PIL conversion failed: {e}")
        
        # Fallback: create a simple placeholder
        win32gui.DeleteObject(hbmp)
        win32gui.DeleteDC(hdc_mem)
        win32gui.ReleaseDC(0, hdc)
        
        # Create simple placeholder
        pixmap = QPixmap(size[0], size[1])
        pixmap.fill(Qt.lightGray)
        painter = QPainter(pixmap)
        painter.setPen(Qt.darkGray)
        painter.drawRect(2, 2, size[0]-4, size[1]-4)
        painter.drawText(pixmap.rect(), Qt.AlignCenter, "APP")
        painter.end()
        
        return pixmap
                
    except Exception as e:
        print(f"Error in icon conversion: {e}")
        
        # Ultimate fallback
        pixmap = QPixmap(size[0], size[1])
        pixmap.fill(Qt.lightGray)
        painter = QPainter(pixmap)
        painter.setPen(Qt.darkGray)
        painter.drawRect(2, 2, size[0]-4, size[1]-4)
        painter.drawText(pixmap.rect(), Qt.AlignCenter, "?")
        painter.end()
        
        return pixmap

user32 = ctypes.windll.user32

target_hwnd = None
target_title = None
stop_flag = False

VK_KEYS = {
    'page up': 0x21,
    'page down': 0x22,
    'left': 0x25,
    'right': 0x27,
    'up': 0x26,
    'down': 0x28
}

def send_key_to_window_advanced(hwnd, vk):
    """Advanced method using AttachThreadInput for stubborn applications"""
    if not hwnd or not win32gui.IsWindow(hwnd):
        return False
        
    try:
        import win32process
        import win32api
        import threading
        
        # Get the thread ID of the current process and target window
        current_thread_id = win32api.GetCurrentThreadId()
        target_thread_id, target_pid = win32process.GetWindowThreadProcessId(hwnd)
        
        # Create extended key scan codes with proper flags
        # bit 24 = extended key flag
        extended_keys = {0x21, 0x22, 0x25, 0x27}  # PAGE_UP, PAGE_DOWN, LEFT, RIGHT
        lparam = 0x00000001  # Standard scancode
        if vk in extended_keys:
            lparam |= (1 << 24)  # Add extended key flag
        
        if current_thread_id != target_thread_id:
            # Attach to the target thread's input queue
            ctypes.windll.user32.AttachThreadInput(current_thread_id, target_thread_id, True)
        
        try:
            # Ensure window is responsive
            if ctypes.windll.user32.IsHungAppWindow(hwnd):
                print(f"Window appears to be hung (not responding)")
                return False
                
            # Send the key using SendInput with proper parameters
            WM_KEYDOWN = 0x0100
            WM_KEYUP = 0x0101
            
            # First ensure window has focus 
            foreground_window = win32gui.GetForegroundWindow()
            
            # Skip if this is the current foreground window to avoid disrupting user
            if foreground_window != hwnd:
                # Try to set focus to window - this might fail for background windows
                # but we'll still attempt key send afterwards
                try:
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.01)  # Small delay to allow window to get focus
                except Exception as focus_error:
                    print(f"Could not set focus: {focus_error}")
            
            # Send keys with hardware scancode information
            win32gui.PostMessage(hwnd, WM_KEYDOWN, vk, lparam)
            time.sleep(0.01)  # Small delay between down and up
            win32gui.PostMessage(hwnd, WM_KEYUP, vk, lparam | 0xC0000000)  # Key up flag
            result = True
        finally:
            # Always detach from the thread input queue
            if current_thread_id != target_thread_id:
                ctypes.windll.user32.AttachThreadInput(current_thread_id, target_thread_id, False)
        
        return result
        
    except Exception as e:
        print(f"Advanced method failed: {e}")
        return False


def send_key_to_window(hwnd, vk):
    """Send a key press to a specific window using multiple techniques"""
    if not hwnd or not win32gui.IsWindow(hwnd):
        return False
    
    # Try several methods in order of preference
    methods = [
        "advanced",    # Our best method (with thread attachment)
        "post",        # PostMessage (no focus change)  
        "send",        # SendMessage (blocking)
        "sendinput",   # SendInput (global input injection)
    ]
    
    # Create extended key scan codes with proper flags for certain special keys
    extended_keys = {0x21, 0x22, 0x25, 0x27}  # PAGE_UP, PAGE_DOWN, LEFT, RIGHT
    lparam = 0x00000001  # Standard scancode
    if vk in extended_keys:
        lparam |= (1 << 24)  # Add extended key flag
    
    # Try all methods in sequence until one works
    for method in methods:
        try:
            if method == "advanced":
                if send_key_to_window_advanced(hwnd, vk):
                    print(f"Advanced method succeeded")
                    return True
                
            elif method == "post":
                WM_KEYDOWN = 0x0100
                WM_KEYUP = 0x0101
                win32gui.PostMessage(hwnd, WM_KEYDOWN, vk, lparam)
                time.sleep(0.01)  # Small delay between down and up
                win32gui.PostMessage(hwnd, WM_KEYUP, vk, lparam | 0xC0000000)
                print(f"PostMessage succeeded")
                return True
                
            elif method == "send":
                WM_KEYDOWN = 0x0100
                WM_KEYUP = 0x0101
                win32gui.SendMessage(hwnd, WM_KEYDOWN, vk, lparam)
                win32gui.SendMessage(hwnd, WM_KEYUP, vk, lparam | 0xC0000000)
                print(f"SendMessage succeeded")
                return True
                
            elif method == "sendinput":
                # Use SendInput as last resort - will change system focus
                try:
                    # First set focus to the window
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.05)  # Wait for focus
                    
                    # Then send the key using system keyboard
                    from ctypes import wintypes
                    
                    # Define required structures
                    INPUT_KEYBOARD = 1
                    KEYEVENTF_KEYUP = 0x0002
                    KEYEVENTF_EXTENDEDKEY = 0x0001
                    
                    class KEYBDINPUT(ctypes.Structure):
                        _fields_ = [
                            ("wVk", wintypes.WORD),
                            ("wScan", wintypes.WORD),
                            ("dwFlags", wintypes.DWORD),
                            ("time", wintypes.DWORD),
                            ("dwExtraInfo", wintypes.PULONG)
                        ]
                        
                    class INPUT(ctypes.Structure):
                        _fields_ = [
                            ("type", wintypes.DWORD),
                            ("ki", KEYBDINPUT)
                        ]
                    
                    # Key down
                    input_down = INPUT()
                    input_down.type = INPUT_KEYBOARD
                    input_down.ki.wVk = vk
                    if vk in extended_keys:
                        input_down.ki.dwFlags = KEYEVENTF_EXTENDEDKEY
                    
                    # Key up
                    input_up = INPUT()
                    input_up.type = INPUT_KEYBOARD
                    input_up.ki.wVk = vk
                    if vk in extended_keys:
                        input_up.ki.dwFlags = KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP
                    else:
                        input_up.ki.dwFlags = KEYEVENTF_KEYUP
                    
                    # Send keys
                    inputs = (INPUT * 2)(input_down, input_up)
                    ctypes.windll.user32.SendInput(2, inputs, ctypes.sizeof(INPUT))
                    print(f"SendInput succeeded")
                    return True
                except Exception as si_error:
                    print(f"SendInput failed: {si_error}")
                    
        except Exception as e:
            print(f"{method} method failed: {e}")
    
    print("All key sending methods failed")
    return False


def key_listener():
    global stop_flag, target_hwnd, target_title
    key_pressed = {key: False for key in VK_KEYS.keys()}
    key_last_press_time = {key: 0 for key in VK_KEYS.keys()}
    debounce_time = 0.15  # Debounce time in seconds (adjust as needed)
    failed_attempts = 0
    max_failed_attempts = 5
    
    while not stop_flag:
        if target_hwnd and win32gui.IsWindow(target_hwnd):
            current_time = time.time()
            
            # Check if target window is the active (foreground) window
            foreground_hwnd = win32gui.GetForegroundWindow()
            is_target_focused = (foreground_hwnd == target_hwnd)
            
            for key, vk in VK_KEYS.items():
                is_pressed = keyboard.is_pressed(key)
                
                # Only send key if it's newly pressed (not held down) and debounce time has passed
                if is_pressed and not key_pressed[key]:
                    time_since_last_press = current_time - key_last_press_time[key]
                    
                    # Check if the target window is already focused - if so, don't send keys
                    # because Windows will deliver them directly (avoiding double triggering)
                    if is_target_focused:
                        print(f"Target window is already focused - not redirecting {key}")
                    # Otherwise check debounce and send keys if needed
                    elif time_since_last_press > debounce_time:
                        print(f"Sending {key} to {target_title} (time since last: {time_since_last_press:.2f}s)")
                        success = send_key_to_window(target_hwnd, vk)
                        
                        # Update last press time
                        key_last_press_time[key] = current_time
                        
                        if success:
                            failed_attempts = 0  # Reset counter on success
                        else:
                            failed_attempts += 1
                            if failed_attempts >= max_failed_attempts:
                                print(f"Warning: Failed to send keys {failed_attempts} times. Target window may not be responding.")
                                failed_attempts = 0  # Reset to avoid spam
                    else:
                        print(f"Ignoring {key} press - too soon after last press ({time_since_last_press:.2f}s < {debounce_time}s)")
                    
                    key_pressed[key] = True
                elif not is_pressed:
                    key_pressed[key] = False
        elif target_hwnd and not win32gui.IsWindow(target_hwnd):
            # Target window no longer exists
            print(f"Target window '{target_title}' is no longer available")
            target_hwnd = None
            target_title = None
                    
        time.sleep(0.05)


def get_window_thumbnail(hwnd, size=(150, 100)):
    """Capture a thumbnail of the specified window using multiple approaches"""
    if not THUMBNAIL_SUPPORT:
        return get_window_thumbnail_fallback(hwnd, size)
        
    try:
        # Get window dimensions and position
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        width = right - left
        height = bottom - top
        
        window_title = win32gui.GetWindowText(hwnd)
        print(f"Capturing thumbnail for window: {window_title}, size: {width}x{height}")
        
        if width <= 0 or height <= 0 or width > 3000 or height > 2000:
            print(f"Window has invalid dimensions: {width}x{height}")
            return get_window_thumbnail_fallback(hwnd, size)
        
        # First try Windows PrintWindow API - most accurate for capturing exact window
        try:
            # Create device contexts and bitmap
            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()
            
            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
            saveDC.SelectObject(saveBitMap)
            
            # Capture window with PrintWindow - using PW_RENDERFULLCONTENT flag for best results
            ctypes.windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 2)  # 2 = PW_RENDERFULLCONTENT
            
            # Convert to PIL Image - use in-memory bitmap approach for best orientation handling
            from io import BytesIO
            import struct
            
            # Get raw bitmap data
            bmpstr = saveBitMap.GetBitmapBits(True)
            bmpinfo = saveBitMap.GetInfo()
            
            # Create in-memory BMP file for proper processing - this handles orientation correctly
            bmp_file = BytesIO()
            
            # BMP header (14 bytes)
            bmp_file.write(b'BM')  # Signature
            file_size = 14 + 40 + len(bmpstr)  # Header + DIB header + bitmap data
            bmp_file.write(struct.pack('<I', file_size))  # FileSize
            bmp_file.write(struct.pack('<I', 0))  # Reserved
            bmp_file.write(struct.pack('<I', 14 + 40))  # DataOffset
            
            # DIB header (40 bytes) - using negative height for top-down orientation
            bmp_file.write(struct.pack('<I', 40))  # HeaderSize
            bmp_file.write(struct.pack('<i', bmpinfo['bmWidth']))  # Width
            bmp_file.write(struct.pack('<i', -bmpinfo['bmHeight']))  # Height (negative = top-down)
            bmp_file.write(struct.pack('<H', 1))  # Planes
            bmp_file.write(struct.pack('<H', 32))  # BitCount
            bmp_file.write(struct.pack('<I', 0))  # Compression
            bmp_file.write(struct.pack('<I', len(bmpstr)))  # ImageSize
            bmp_file.write(struct.pack('<i', 0))  # XPelsPerMeter
            bmp_file.write(struct.pack('<i', 0))  # YPelsPerMeter
            bmp_file.write(struct.pack('<I', 0))  # ClrUsed
            bmp_file.write(struct.pack('<I', 0))  # ClrImportant
            
            # Bitmap data
            bmp_file.write(bmpstr)
            bmp_file.seek(0)
            
            # Load from in-memory BMP file
            img = Image.open(bmp_file)
            
            # Resize to thumbnail
            img.thumbnail(size, Image.Resampling.LANCZOS)
            
            # Convert to QPixmap
            pixmap = pil_to_qpixmap(img)
            
            # Clean up resources
            win32gui.DeleteObject(saveBitMap.GetHandle())
            saveDC.DeleteDC()
            mfcDC.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwndDC)
            
            if pixmap and not pixmap.isNull():
                print(f"Successfully captured window with PrintWindow API")
                return pixmap
        except Exception as pw_error:
            print(f"PrintWindow capture failed: {pw_error}")
        
        # Try alternate approach with Qt's window capture
        try:
            from PyQt5.QtWidgets import QApplication
            from PyQt5.QtCore import QRect
            
            # Get the screen containing this window
            screens = QApplication.screens()
            
            # Create rectangle for the window
            window_rect = QRect(left, top, width, height)
            
            # Check which screen contains most of this window
            containing_screen = None
            max_overlap = 0
            
            for screen in screens:
                screen_geom = screen.geometry()
                # Calculate intersection between window and this screen
                overlap_rect = window_rect.intersected(screen_geom)
                overlap_area = overlap_rect.width() * overlap_rect.height()
                
                if overlap_area > max_overlap:
                    max_overlap = overlap_area
                    containing_screen = screen
            
            if not containing_screen:
                # Fallback to primary screen
                containing_screen = QApplication.primaryScreen()
                
            # Try with direct window handle capture first
            screenshot = None
            try:
                # Use window handle directly - this should only get the window content
                screenshot = containing_screen.grabWindow(
                    hwnd,  # Window handle
                    0, 0,  # Start at origin
                    width, # Width
                    height # Height
                )
            except Exception:
                # Fall back to screen area capture as last resort
                screenshot = containing_screen.grabWindow(
                    0,    # Desktop/screen
                    left, # X 
                    top,  # Y
                    width,# Width
                    height# Height
                )
            
            if not screenshot.isNull():
                print("Successfully captured screenshot using Qt screen grabbing")
                
                # Scale to thumbnail size
                thumbnail = screenshot.scaled(size[0], size[1], Qt.KeepAspectRatio, Qt.SmoothTransformation)
                return thumbnail
                
        except Exception as qt_error:
            print(f"Qt screenshot method failed: {qt_error}")
            
        # Fallback to modified Windows API method as alternative
        try:
            # Use a different approach with Win32 API
            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()
            
            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
            saveDC.SelectObject(saveBitMap)
            
            # Try PrintWindow with PW_RENDERFULLCONTENT flag (Windows 8.1+)
            result = ctypes.windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 2)  # PW_RENDERFULLCONTENT
            
            if result:
                # Convert the bitmap to a format we can work with
                # Using a modern approach with DIB sections to avoid bitmap orientation issues
                
                # First get bitmap object 
                bitmap_handle = saveBitMap.GetHandle()
                
                # Create DIB (device-independent bitmap) info
                import struct
                
                class BITMAPINFOHEADER(ctypes.Structure):
                    _fields_ = [
                        ("biSize", ctypes.c_uint32),
                        ("biWidth", ctypes.c_int),
                        ("biHeight", ctypes.c_int),
                        ("biPlanes", ctypes.c_uint16),
                        ("biBitCount", ctypes.c_uint16),
                        ("biCompression", ctypes.c_uint32),
                        ("biSizeImage", ctypes.c_uint32),
                        ("biXPelsPerMeter", ctypes.c_int),
                        ("biYPelsPerMeter", ctypes.c_int),
                        ("biClrUsed", ctypes.c_uint32),
                        ("biClrImportant", ctypes.c_uint32)
                    ]
                
                # Create header with TOP-DOWN orientation (negative height = top-down)
                header = BITMAPINFOHEADER()
                header.biSize = ctypes.sizeof(BITMAPINFOHEADER)
                header.biWidth = width
                header.biHeight = -height  # Negative for TOP-DOWN orientation
                header.biPlanes = 1
                header.biBitCount = 32
                header.biCompression = 0  # BI_RGB
                
                # Get the actual bitmap data
                bits = saveBitMap.GetBitmapBits(True)
                
                # Create PIL image directly from the raw data using a more reliable approach
                from io import BytesIO
                
                # Create PIL image from bitmap data
                img = Image.frombuffer(
                    'RGBA', 
                    (width, height),
                    bits, 
                    'raw', 
                    'BGRA', 
                    0, 
                    1
                )
                
                # Windows bitmaps are stored bottom-up by default, so no need to flip
                # If images appear upside down, we would uncomment the following line
                # img = img.transpose(Image.FLIP_TOP_BOTTOM)
                
                # Open with PIL directly from memory
                try:
                    
                    # Resize to thumbnail size
                    img.thumbnail(size, Image.Resampling.LANCZOS)
                    
                    # Convert to QPixmap
                    pixmap = pil_to_qpixmap(img)
                    
                    if pixmap and not pixmap.isNull():
                        print("Successfully captured using DIB method")
                        
                        # Clean up
                        win32gui.DeleteObject(saveBitMap.GetHandle())
                        saveDC.DeleteDC()
                        mfcDC.DeleteDC()
                        win32gui.ReleaseDC(hwnd, hwndDC)
                        
                        return pixmap
                except Exception as e:
                    print(f"DIB conversion failed: {e}")
            
            # Clean up resources
            win32gui.DeleteObject(saveBitMap.GetHandle())
            saveDC.DeleteDC()
            mfcDC.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwndDC)
            
        except Exception as e:
            print(f"Win32 bitmap creation failed: {e}")
        
        # If we get here, all direct screen capture methods have failed
        return get_window_thumbnail_fallback(hwnd, size)
            
    except Exception as e:
        print(f"Error in window thumbnail capture: {e}")
        return get_window_thumbnail_fallback(hwnd, size)


def get_window_thumbnail_fallback(hwnd, size=(150, 100)):
    """Fallback method using basic window icon"""
    try:
        print(f"Using fallback method for: {win32gui.GetWindowText(hwnd)}")
        
        # Try multiple ways to get an icon
        icon_handle = None
        
        # Method 1: Try to get large icon
        try:
            icon_handle = win32gui.SendMessage(hwnd, win32con.WM_GETICON, win32con.ICON_BIG, 0)
            if icon_handle:
                print("Got large icon")
        except Exception as e:
            print(f"Failed to get large icon: {e}")
        
        # Method 2: Try to get small icon
        if not icon_handle:
            try:
                icon_handle = win32gui.SendMessage(hwnd, win32con.WM_GETICON, win32con.ICON_SMALL, 0)
                if icon_handle:
                    print("Got small icon")
            except Exception as e:
                print(f"Failed to get small icon: {e}")
        
        # Method 3: Try class icon
        if not icon_handle:
            try:
                icon_handle = win32gui.GetClassLong(hwnd, win32con.GCL_HICON)
                if icon_handle:
                    print("Got class large icon")
            except Exception as e:
                print(f"Failed to get class large icon: {e}")
        
        # Method 4: Try small class icon
        if not icon_handle:
            try:
                icon_handle = win32gui.GetClassLong(hwnd, win32con.GCL_HICONSM)
                if icon_handle:
                    print("Got class small icon")
            except Exception as e:
                print(f"Failed to get class small icon: {e}")
        
        # Method 5: Try to get the executable icon
        if not icon_handle:
            try:
                import win32process
                import win32api
                
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
                exe_path = win32process.GetModuleFileNameEx(handle, 0)
                win32api.CloseHandle(handle)
                
                # Extract icon from executable
                large_icons, small_icons = win32gui.ExtractIconEx(exe_path, 0)
                if large_icons:
                    icon_handle = large_icons[0]
                    print("Got executable icon")
                elif small_icons:
                    icon_handle = small_icons[0]
                    print("Got executable small icon")
                    
            except Exception as e:
                print(f"Failed to get executable icon: {e}")
        
        if icon_handle:
            # Convert to QPixmap using our custom function
            try:
                pixmap = convert_hicon_to_qpixmap(icon_handle, size)
                if pixmap and not pixmap.isNull():
                    print(f"Successfully created icon pixmap: {pixmap.width()}x{pixmap.height()}")
                    return pixmap
            except Exception as e:
                print(f"Failed to convert icon to pixmap: {e}")
        
        # Ultimate fallback: create a simple placeholder
        print("Creating simple placeholder")
        pixmap = QPixmap(size[0], size[1])
        pixmap.fill(Qt.lightGray)
        painter = QPainter(pixmap)
        painter.setPen(Qt.darkGray)
        painter.drawRect(2, 2, size[0]-4, size[1]-4)
        
        # Try to get just the window title for the placeholder
        try:
            title = win32gui.GetWindowText(hwnd)
            if title:
                # Use first few characters of title
                short_title = title[:8] if len(title) > 8 else title
                painter.drawText(pixmap.rect(), Qt.AlignCenter, short_title)
            else:
                painter.drawText(pixmap.rect(), Qt.AlignCenter, "WIN")
        except:
            painter.drawText(pixmap.rect(), Qt.AlignCenter, "APP")
        
        painter.end()
        return pixmap
                
    except Exception as e:
        print(f"Error in fallback method: {e}")
    
    # Last resort fallback
    print("Using last resort fallback")
    pixmap = QPixmap(size[0], size[1])
    pixmap.fill(Qt.red)
    painter = QPainter(pixmap)
    painter.setPen(Qt.white)
    painter.drawText(pixmap.rect(), Qt.AlignCenter, "ERR")
    painter.end()
    return pixmap


def get_open_windows():
    windows = []

    def enum_handler(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title:
                windows.append((hwnd, title))

    win32gui.EnumWindows(enum_handler, None)
    return windows


class WindowListItem(QWidget):
    """Custom widget for displaying window with thumbnail"""
    def __init__(self, hwnd, title, thumbnail=None):
        super().__init__()
        self.hwnd = hwnd
        self.title = title
        
        layout = QHBoxLayout()
        layout.setContentsMargins(5, 5, 5, 5)
        
        # Thumbnail label
        self.thumbnail_label = QLabel()
        self.thumbnail_label.setFixedSize(80, 60)
        self.thumbnail_label.setStyleSheet("border: 1px solid gray; background-color: #f0f0f0;")
        self.thumbnail_label.setAlignment(Qt.AlignCenter)
        
        if thumbnail and not thumbnail.isNull():
            scaled_thumbnail = thumbnail.scaled(78, 58, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.thumbnail_label.setPixmap(scaled_thumbnail)
        else:
            self.thumbnail_label.setText("No\nPreview")
            self.thumbnail_label.setStyleSheet(self.thumbnail_label.styleSheet() + "color: gray; font-size: 10px;")
        
        # Window info label
        info_label = QLabel(f"<b>{title}</b><br><small>HWND: {hwnd}</small>")
        info_label.setWordWrap(True)
        info_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        
        layout.addWidget(self.thumbnail_label)
        layout.addWidget(info_label, 1)
        
        self.setLayout(layout)


class WindowSelector(QWidget):
    def __init__(self, parent_app=None):
        super().__init__()
        self.parent_app = parent_app
        self.setWindowTitle("Select Target Window")
        self.setGeometry(100, 100, 600, 500)

        self.layout = QVBoxLayout()
        self.list_widget = QListWidget()
        self.list_widget.setAlternatingRowColors(True)
        
        # Add loading label
        self.loading_label = QLabel("Loading windows and capturing thumbnails...")
        self.loading_label.setAlignment(Qt.AlignCenter)
        self.loading_label.setStyleSheet("font-size: 12px; color: #666; padding: 10px;")
        
        self.refresh_btn = QPushButton("Refresh List")
        self.set_target_btn = QPushButton("Set as Target")

        self.layout.addWidget(self.loading_label)
        self.layout.addWidget(self.list_widget)
        self.layout.addWidget(self.refresh_btn)
        self.layout.addWidget(self.set_target_btn)

        self.setLayout(self.layout)

        self.refresh_btn.clicked.connect(self.load_windows)
        self.set_target_btn.clicked.connect(self.set_target)
        
        # Load windows after a short delay to show the loading message
        QTimer.singleShot(100, self.load_windows)

    def load_windows(self):
        self.loading_label.show()
        self.list_widget.hide()
        self.list_widget.clear()
        self.windows = get_open_windows()
        
        # Process windows and create thumbnails
        for i, (hwnd, title) in enumerate(self.windows):
            # Update loading message
            self.loading_label.setText(f"Loading windows... ({i+1}/{len(self.windows)})")
            QApplication.processEvents()  # Allow GUI to update
            
            # Try to get thumbnail
            thumbnail = get_window_thumbnail(hwnd)
            if not thumbnail:
                thumbnail = get_window_thumbnail_fallback(hwnd)
            
            # Create custom widget
            window_widget = WindowListItem(hwnd, title, thumbnail)
            
            # Create list item
            item = QListWidgetItem()
            item.setSizeHint(QSize(580, 80))  # Set item size
            
            # Add to list
            self.list_widget.addItem(item)
            self.list_widget.setItemWidget(item, window_widget)
        
        # Hide loading message and show list
        self.loading_label.hide()
        self.list_widget.show()

    def set_target(self):
        global target_hwnd, target_title
        index = self.list_widget.currentRow()
        if index >= 0:
            target_hwnd, target_title = self.windows[index]
            if self.parent_app:
                self.parent_app.update_tooltip()
            QMessageBox.information(self, "Target Set", f"Target window set:\n{target_title}\n\nYour arrow keys and page up/down will now be redirected to this window.")
            self.close()
        else:
            QMessageBox.warning(self, "No Selection", "Please select a window first.")

def is_windows_dark_mode():
    """Check if Windows is in dark mode"""
    try:
        # Open the key where Windows stores its theme settings
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, 
                           r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
        
        # Value 0 means dark mode, 1 means light mode
        value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        return value == 0
    except Exception as e:
        print(f"Error checking dark mode: {e}")
        return False


class TrayApp(QApplication):
    def __init__(self, sys_argv):
        super().__init__(sys_argv)

        # Prevent app from closing when last window closes
        self.setQuitOnLastWindowClosed(False)

        # Set up system tray
        try:
            if is_windows_dark_mode():
                # Use dark icon if system is in dark mode
                icon = QIcon("icon_dark.png")
            else:
                # Use light icon for normal mode
                icon = QIcon("icon.png")
            if icon.isNull():
                # Create a simple default icon if file doesn't exist
                icon = self.style().standardIcon(self.style().SP_ComputerIcon)
        except:
            icon = self.style().standardIcon(self.style().SP_ComputerIcon)
            
        self.tray = QSystemTrayIcon(icon)
        self.menu = QMenu()

        self.select_action = QAction("Select Window")
        self.show_target_action = QAction("Show Current Target")
        self.run_at_startup_action = QAction("Run at Startup")
        self.run_at_startup_action.setCheckable(True)
        self.run_at_startup_action.setChecked(self.is_set_to_run_at_startup())
        self.quit_action = QAction("Quit")

        self.menu.addAction(self.select_action)
        self.menu.addAction(self.show_target_action)
        self.menu.addSeparator()
        self.menu.addAction(self.run_at_startup_action)
        self.menu.addSeparator()
        self.menu.addAction(self.quit_action)

        self.tray.setContextMenu(self.menu)
        self.tray.setToolTip("PPT Redirector - No target selected")
        
        # Check if system tray is available
        if not QSystemTrayIcon.isSystemTrayAvailable():
            QMessageBox.critical(None, "System Tray", "System tray is not available on this system.")
            return
            
        self.tray.show()

        self.select_action.triggered.connect(self.open_selector)
        self.show_target_action.triggered.connect(self.show_current_target)
        self.run_at_startup_action.triggered.connect(self.toggle_run_at_startup)
        self.quit_action.triggered.connect(self.quit_all)
        
        # Allow single-click or double-click on tray icon to open selector
        self.tray.activated.connect(self.tray_icon_activated)

        # Start key listener in background
        self.listener_thread = threading.Thread(target=key_listener, daemon=True)
        self.listener_thread.start()
        
        print("PPT Redirector started. Click or right-click the system tray icon to select a target window.")

    def tray_icon_activated(self, reason):
        """Handle tray icon activation (single click, double-click, etc.)"""
        if reason == QSystemTrayIcon.DoubleClick or reason == QSystemTrayIcon.Trigger:
            # Trigger corresponds to a normal left click
            self.open_selector()

    def open_selector(self):
        self.selector = WindowSelector(parent_app=self)
        self.selector.show()

    def show_current_target(self):
        global target_hwnd, target_title
        if target_hwnd and target_title:
            if win32gui.IsWindow(target_hwnd):
                QMessageBox.information(None, "Current Target", f"Currently redirecting keys to:\n{target_title}")
            else:
                QMessageBox.warning(None, "Target Lost", "The target window is no longer available. Please select a new target.")
                target_hwnd = None
                target_title = None
                self.tray.setToolTip("PPT Redirector - No target selected")
        else:
            QMessageBox.information(None, "No Target", "No target window selected. Use 'Select Window' to choose a target.")

    def is_set_to_run_at_startup(self):
        """Check if the application is set to run at startup"""
        try:
            startup_folder = os.path.join(os.environ['APPDATA'], r'Microsoft\Windows\Start Menu\Programs\Startup')
            shortcut_path = os.path.join(startup_folder, "PPT Redirector.lnk")
            return os.path.exists(shortcut_path)
        except Exception as e:
            print(f"Error checking startup status: {e}")
            return False

    def toggle_run_at_startup(self):
        """Toggle whether the application runs at startup"""
        try:
            # Get paths
            script_dir = os.path.dirname(os.path.abspath(__file__))
            batch_file_path = os.path.join(script_dir, "start_ppt_redirector.bat")
            startup_folder = os.path.join(os.environ['APPDATA'], r'Microsoft\Windows\Start Menu\Programs\Startup')
            shortcut_path = os.path.join(startup_folder, "PPT Redirector.lnk")
            
            # Check current state
            is_enabled = os.path.exists(shortcut_path)
            
            if is_enabled:
                # Remove from startup
                if os.path.exists(shortcut_path):
                    os.remove(shortcut_path)
                self.run_at_startup_action.setChecked(False)
                self.tray.showMessage("PPT Redirector", "Removed from Windows startup", QSystemTrayIcon.Information, 2000)
            else:
                # Add to startup by creating a shortcut
                try:
                    # First ensure the batch file exists with updated content
                    with open(batch_file_path, 'w') as f:
                        f.write('@echo off\ncd /d "%~dp0"\nstart /b "" pythonw main.py\nexit')
                    
                    # Create the shortcut using PowerShell - direct link to pythonw for completely hidden execution
                    main_script_path = os.path.join(script_dir, "main.py")
                    powershell_cmd = f'''
                    $WScriptShell = New-Object -ComObject WScript.Shell
                    $Shortcut = $WScriptShell.CreateShortcut('{shortcut_path}')
                    $Shortcut.TargetPath = 'pythonw.exe'
                    $Shortcut.Arguments = '"{main_script_path}"'
                    $Shortcut.Description = 'Start PPT Redirector'
                    $Shortcut.WorkingDirectory = '{script_dir}'
                    $Shortcut.WindowStyle = 0
                    $Shortcut.Save()
                    '''
                    
                    subprocess.run(['powershell', '-Command', powershell_cmd], 
                                   capture_output=True, text=True, check=True)
                    
                    # Check if the shortcut was created
                    if os.path.exists(shortcut_path):
                        self.run_at_startup_action.setChecked(True)
                        self.tray.showMessage("PPT Redirector", "Added to Windows startup", QSystemTrayIcon.Information, 2000)
                    else:
                        raise Exception("Shortcut file was not created")
                        
                except Exception as e:
                    print(f"Error creating startup shortcut: {e}")
                    QMessageBox.warning(None, "Startup Configuration", 
                                        f"Could not add to startup: {e}\n\nTry running as administrator or use add_to_startup.ps1 script.")
                    self.run_at_startup_action.setChecked(False)
        except Exception as e:
            print(f"Error toggling startup: {e}")
            QMessageBox.warning(None, "Startup Error", f"Error managing startup settings: {e}")
            # Refresh checkbox state
            self.run_at_startup_action.setChecked(self.is_set_to_run_at_startup())
    
    def update_tooltip(self):
        """Update the system tray tooltip based on current target"""
        global target_title
        if target_title:
            self.tray.setToolTip(f"PPT Redirector - Target: {target_title}")
        else:
            self.tray.setToolTip("PPT Redirector - No target selected")

    def quit_all(self):
        global stop_flag
        stop_flag = True
        self.tray.hide()
        self.quit()


def check_for_updates():
    """Check for updates from GitHub repository"""
    try:
        # Get the current directory (where main.py is located)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        
        print("Checking for updates from GitHub repository...")
        update_found = False
        update_message = ""
        
        # Check if it's a git repository
        git_dir = os.path.join(script_dir, ".git")
        if not os.path.exists(git_dir):
            print("Not a git repository. Cloning from GitHub...")
            # Clone the repository if it doesn't exist
            subprocess.run(
                ["git", "clone", "https://github.com/rol4400/PPT-Focuser.git", "."],
                cwd=script_dir,
                check=True
            )
            update_found = True
            update_message = "Repository cloned successfully."
        else:
            # Pull updates if it is a git repository
            result = subprocess.run(
                ["git", "pull", "https://github.com/rol4400/PPT-Focuser.git"],
                cwd=script_dir,
                check=True,
                capture_output=True,
                text=True
            )
            
            if "Already up to date" in result.stdout:
                print("Application is up to date.")
            else:
                print(f"Updates pulled from GitHub: {result.stdout.strip()}")
                print("Restart might be required for changes to take effect.")
                update_found = True
                update_message = f"Updates installed: {result.stdout.strip()}"
        
        return update_found, update_message
    except Exception as e:
        print(f"Error checking for updates: {e}")
        return False, str(e)


def main():
    # Start the application first to access Qt functionality
    app = TrayApp(sys.argv)
    
    # Check for updates after application is initialized
    def delayed_update_check():
        # Wait a short moment to ensure tray icon is ready
        time.sleep(1)
        update_found, update_message = check_for_updates()
        
        # Show notification if update was found
        if update_found and hasattr(app, 'tray'):
            app.tray.showMessage(
                "PPT Redirector Updated",
                update_message + "\nSome changes may require a restart.",
                app.style().standardIcon(app.style().SP_DialogInformationIcon),
                5000  # Show for 5 seconds
            )
    
    # Run the update check in a separate thread to avoid blocking the UI
    update_thread = threading.Thread(target=delayed_update_check, daemon=True)
    update_thread.start()
    
    # Start the application event loop
    sys.exit(app.exec_())


if __name__ == "__main__":
    if "--show-keys" in sys.argv:
        print("Key detection mode: Press any key on your PPT clicker (Press ESC to exit).")

        def on_key_event(event):
            print(f"Key pressed: {event.name} (scan_code={event.scan_code}, vk_code={event.scan_code})")
            if event.name == 'esc':
                print("Exiting key detection.")
                keyboard.unhook_all()
                sys.exit(0)

        keyboard.on_press(on_key_event)
        while True:
            time.sleep(0.1)
    else:
        main()
