# pywinctl/_win_api.py

import sys
import time

if sys.platform != 'win32':
    raise ImportError("pywinctl requires the Windows operating system.")

try:
    import win32gui
    import win32con
    import win32api
    import win32process
    import pywintypes
except ImportError:
    raise ImportError("pywin32 library is required. Please install it using 'pip install pywin32'.")

try:
    import psutil
except ImportError:
    psutil = None # psutil is optional but recommended for process info

from ._exceptions import InvalidWindowError, WindowsAPIError

# --- Constants ---
HWND_TOPMOST = -1
HWND_NOTOPMOST = -2
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
SWP_NOZORDER = 0x0004
SWP_SHOWWINDOW = 0x0040
SWP_NOACTIVATE = 0x0010


# --- Helper Functions ---

def _check_hwnd(hwnd):
    """Check if the window handle is valid."""
    if not isinstance(hwnd, int) or hwnd == 0:
        raise InvalidWindowError(f"Invalid HWND provided: {hwnd}")
    if not win32gui.IsWindow(hwnd):
        raise InvalidWindowError(f"Window with HWND {hwnd} does not exist or is invalid.")

def _get_last_error():
    """Get the last error code from Windows API."""
    return win32api.GetLastError()

def _raise_win_api_error(message, hwnd=None):
    """Raise a WindowsAPIError with the last error code."""
    error_code = _get_last_error()
    if hwnd is not None:
        message = f"{message} (HWND: {hwnd})"
    raise WindowsAPIError(message, error_code)


# --- Core API Wrappers ---

def set_window_pos(hwnd: int, x: int, y: int, width: int, height: int):
    """Move and resize a window using SetWindowPos."""
    _check_hwnd(hwnd)
    try:
        # Flags: Don't change Z-order, don't activate
        flags = SWP_NOZORDER | SWP_NOACTIVATE
        win32gui.SetWindowPos(hwnd, 0, x, y, width, height, flags)
    except pywintypes.error as e:
        # Error code 1400: Invalid window handle (often means window closed)
        if e.winerror == 1400:
            raise InvalidWindowError(f"Window with HWND {hwnd} is invalid (possibly closed).") from e
        _raise_win_api_error(f"Failed to set window position/size for HWND {hwnd}", hwnd)

def move_window(hwnd: int, x: int, y: int):
    """Move a window without changing its size or Z-order."""
    _check_hwnd(hwnd)
    try:
        # Flags: Don't change size, don't change Z-order, don't activate
        flags = SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE
        win32gui.SetWindowPos(hwnd, 0, x, y, 0, 0, flags)
    except pywintypes.error as e:
        if e.winerror == 1400:
            raise InvalidWindowError(f"Window with HWND {hwnd} is invalid.") from e
        _raise_win_api_error(f"Failed to move window for HWND {hwnd}", hwnd)

def resize_window(hwnd: int, width: int, height: int):
    """Resize a window without moving it or changing its Z-order."""
    _check_hwnd(hwnd)
    try:
        # Flags: Don't move, don't change Z-order, don't activate
        flags = SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE
        win32gui.SetWindowPos(hwnd, 0, 0, 0, width, height, flags)
    except pywintypes.error as e:
        if e.winerror == 1400:
            raise InvalidWindowError(f"Window with HWND {hwnd} is invalid.") from e
        _raise_win_api_error(f"Failed to resize window for HWND {hwnd}", hwnd)

def show_window(hwnd: int, command: int):
    """Show, hide, minimize, maximize, or restore a window."""
    _check_hwnd(hwnd)
    try:
        # Use ShowWindowAsync for potentially better responsiveness with unresponsive apps
        # Fallback to ShowWindow if ShowWindowAsync isn't readily available or fails
        try:
            win32gui.ShowWindowAsync(hwnd, command)
        except AttributeError: # If ShowWindowAsync isn't in the pywin32 version
             if not win32gui.ShowWindow(hwnd, command):
                 # ShowWindow returns 0 if it was previously hidden, non-zero otherwise
                 # Check IsWindowVisible to confirm success, especially for SW_SHOW
                 if command == win32con.SW_SHOW and not win32gui.IsWindowVisible(hwnd):
                     time.sleep(0.1) # Give it a moment
                     if not win32gui.IsWindowVisible(hwnd):
                          _raise_win_api_error(f"Failed to show window command {command} for HWND {hwnd}", hwnd)

    except pywintypes.error as e:
         if e.winerror == 1400:
             raise InvalidWindowError(f"Window with HWND {hwnd} is invalid.") from e
         # Error 5: Access Denied (might happen with elevated windows)
         if e.winerror == 5:
             raise WindowsAPIError(f"Access denied trying to control HWND {hwnd}. Try running script as administrator.", e.winerror)
         _raise_win_api_error(f"Failed window command {command} for HWND {hwnd}", hwnd)

def minimize(hwnd: int):
    show_window(hwnd, win32con.SW_MINIMIZE)

def maximize(hwnd: int):
    show_window(hwnd, win32con.SW_MAXIMIZE)

def restore(hwnd: int):
    show_window(hwnd, win32con.SW_RESTORE)

def hide(hwnd: int):
     show_window(hwnd, win32con.SW_HIDE)

def show(hwnd: int):
     show_window(hwnd, win32con.SW_SHOW)


def close_window(hwnd: int):
    """Close a window by sending WM_CLOSE."""
    _check_hwnd(hwnd)
    try:
        win32api.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    except pywintypes.error as e:
        if e.winerror == 1400:
             # Window might already be closed, which is fine for a close operation
             return
        _raise_win_api_error(f"Failed to send WM_CLOSE to HWND {hwnd}", hwnd)


def set_foreground_window(hwnd: int):
    """Bring a window to the foreground."""
    _check_hwnd(hwnd)
    try:
        # Simple SetForegroundWindow often fails due to restrictions.
        # A common workaround:
        # 1. Get the current foreground window's thread ID.
        # 2. Get the target window's thread ID.
        # 3. Attach the input processing mechanism of the two threads.
        # 4. Bring the window to the top (SW_RESTORE if minimized).
        # 5. Set the foreground window.
        # 6. Detach the thread inputs.

        target_thread_id, _ = win32process.GetWindowThreadProcessId(hwnd)
        current_foreground_hwnd = win32gui.GetForegroundWindow()

        # Avoid unnecessary operations if already foreground
        if hwnd == current_foreground_hwnd:
            return

        current_thread_id = win32api.GetCurrentThreadId()
        foreground_thread_id, _ = win32process.GetWindowThreadProcessId(current_foreground_hwnd)

        # Attach threads
        win32process.AttachThreadInput(foreground_thread_id, current_thread_id, True)
        win32process.AttachThreadInput(target_thread_id, current_thread_id, True)

        try:
            # Restore if minimized and bring to top
            if win32gui.IsIconic(hwnd):
                show_window(hwnd, win32con.SW_RESTORE)
            else:
                 # Bring window to top without activating immediately
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 0, 0, 0, 0,
                                      SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE)


            # Attempt to set foreground
            win32gui.SetForegroundWindow(hwnd)
            # Check if successful after a short delay
            time.sleep(0.05)
            if win32gui.GetForegroundWindow() != hwnd:
                # Fallback: Try bringing to top again, might work in some cases
                 win32gui.BringWindowToTop(hwnd)
                 time.sleep(0.05)
                 # Final check
                 if win32gui.GetForegroundWindow() != hwnd:
                      print(f"Warning: Failed to reliably set HWND {hwnd} to foreground.")


        finally:
            # Detach threads
            win32process.AttachThreadInput(target_thread_id, current_thread_id, False)
            win32process.AttachThreadInput(foreground_thread_id, current_thread_id, False)

    except pywintypes.error as e:
        if e.winerror == 1400:
            raise InvalidWindowError(f"Window with HWND {hwnd} is invalid.") from e
        if e.winerror == 5: # Access Denied
             raise WindowsAPIError(f"Access denied trying to set foreground HWND {hwnd}. Try running as administrator.", e.winerror)
        _raise_win_api_error(f"Failed to set foreground window for HWND {hwnd}", hwnd)
    except Exception as e:
        # Catch potential errors during thread attachment/detachment
        raise WindowsAPIError(f"An unexpected error occurred during set_foreground_window for HWND {hwnd}: {e}")


def set_window_title(hwnd: int, title: str):
    """Set the title text of a window."""
    _check_hwnd(hwnd)
    try:
        result = win32gui.SetWindowText(hwnd, title)
        if result == 0: # SetWindowText returns 0 on failure
             last_error = _get_last_error()
             # Error 1400 likely means window closed between check and call
             if last_error != 0 and last_error != 1400:
                 raise WindowsAPIError(f"Failed to set window title for HWND {hwnd}", last_error)
             elif last_error == 1400:
                 raise InvalidWindowError(f"Window with HWND {hwnd} became invalid during title set.")
    except pywintypes.error as e:
        if e.winerror == 1400:
            raise InvalidWindowError(f"Window with HWND {hwnd} is invalid.") from e
        _raise_win_api_error(f"Failed to set window title for HWND {hwnd}", hwnd)

def get_window_title(hwnd: int) -> str:
     """Get the title text of a window."""
     _check_hwnd(hwnd)
     try:
         return win32gui.GetWindowText(hwnd)
     except pywintypes.error as e:
        if e.winerror == 1400:
             raise InvalidWindowError(f"Window with HWND {hwnd} is invalid.") from e
        _raise_win_api_error(f"Failed to get window title for HWND {hwnd}", hwnd)

def get_window_rect(hwnd: int) -> tuple[int, int, int, int]:
    """Get the window's bounding rectangle (left, top, right, bottom)."""
    _check_hwnd(hwnd)
    try:
        return win32gui.GetWindowRect(hwnd)
    except pywintypes.error as e:
        if e.winerror == 1400:
             raise InvalidWindowError(f"Window with HWND {hwnd} is invalid.") from e
        _raise_win_api_error(f"Failed to get window rect for HWND {hwnd}", hwnd)

def get_window_thread_process_id(hwnd: int) -> tuple[int, int]:
    """Get the thread ID and process ID (PID) of the window's creator."""
    _check_hwnd(hwnd)
    try:
        return win32process.GetWindowThreadProcessId(hwnd)
    except pywintypes.error as e:
        if e.winerror == 1400:
            raise InvalidWindowError(f"Window with HWND {hwnd} is invalid.") from e
        _raise_win_api_error(f"Failed to get thread/process ID for HWND {hwnd}", hwnd)

def get_process_info(pid: int) -> dict:
    """Get information about a process using psutil (if available)."""
    if psutil is None:
        return {"error": "psutil library not installed"}
    try:
        proc = psutil.Process(pid)
        return {
            "pid": proc.pid,
            "name": proc.name(),
            "exe": proc.exe(),
            "cwd": proc.cwd(),
            "username": proc.username(),
            "create_time": proc.create_time(),
            "status": proc.status(),
        }
    except psutil.NoSuchProcess:
        return {"error": f"Process with PID {pid} not found"}
    except psutil.AccessDenied:
        return {"error": f"Access denied to process PID {pid}"}
    except Exception as e:
        return {"error": f"Failed to get info for PID {pid}: {e}"}


def set_always_on_top(hwnd: int, enable: bool = True):
    """Set the window's always-on-top status."""
    _check_hwnd(hwnd)
    try:
        z_order = HWND_TOPMOST if enable else HWND_NOTOPMOST
        # Flags: Keep current position and size
        flags = SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE
        win32gui.SetWindowPos(hwnd, z_order, 0, 0, 0, 0, flags)
    except pywintypes.error as e:
        if e.winerror == 1400:
            raise InvalidWindowError(f"Window with HWND {hwnd} is invalid.") from e
        if e.winerror == 5: # Access Denied
             raise WindowsAPIError(f"Access denied trying to set always-on-top for HWND {hwnd}. Try running as administrator.", e.winerror)
        _raise_win_api_error(f"Failed to set always-on-top for HWND {hwnd}", hwnd)

def is_window_visible(hwnd: int) -> bool:
    """Check if the window is visible."""
    _check_hwnd(hwnd)
    try:
        return bool(win32gui.IsWindowVisible(hwnd))
    except pywintypes.error as e:
         if e.winerror == 1400:
             # If handle is invalid, it's not visible
             return False
         _raise_win_api_error(f"Failed to check visibility for HWND {hwnd}", hwnd)

def is_minimized(hwnd: int) -> bool:
    """Check if the window is minimized (iconic)."""
    _check_hwnd(hwnd)
    try:
        return bool(win32gui.IsIconic(hwnd))
    except pywintypes.error as e:
         if e.winerror == 1400:
             # If handle is invalid, consider it not minimized in a functional sense
             return False
         _raise_win_api_error(f"Failed to check minimized state for HWND {hwnd}", hwnd)

def is_maximized(hwnd: int) -> bool:
    """Check if the window is maximized."""
    _check_hwnd(hwnd)
    try:
        return bool(win32gui.IsZoomed(hwnd))
    except pywintypes.error as e:
         if e.winerror == 1400:
             return False
         _raise_win_api_error(f"Failed to check maximized state for HWND {hwnd}", hwnd)

def get_active_window_hwnd() -> int:
    """Get the HWND of the currently active foreground window."""
    try:
        return win32gui.GetForegroundWindow()
    except pywintypes.error as e:
        _raise_win_api_error("Failed to get active window HWND")

def get_window_classname(hwnd: int) -> str:
    """Get the class name of the window."""
    _check_hwnd(hwnd)
    try:
        return win32gui.GetClassName(hwnd)
    except pywintypes.error as e:
        if e.winerror == 1400:
            raise InvalidWindowError(f"Window with HWND {hwnd} is invalid.") from e
        _raise_win_api_error(f"Failed to get class name for HWND {hwnd}", hwnd)