# pywinctl/__init__.py

import sys

if sys.platform != 'win32':
    print("Warning: pywinctl is designed for Windows and may not function correctly on other platforms.")
    # Or raise ImportError("pywinctl requires Windows.")

# Import key components to be accessible directly from the package
from ._main import Window, get_window_by_title, get_active_window, get_all_windows, get_win_test
from ._exceptions import PyWinCtlError, WindowNotFoundError, InvalidWindowError, WindowsAPIError

__version__ = "0.1.0"

# Define what `from pywinctl import *` imports
__all__ = [
    'Window',
    'get_window_by_title',
    'get_active_window',
    'get_all_windows',
    'PyWinCtlError',
    'WindowNotFoundError',
    'InvalidWindowError',
    'WindowsAPIError',
    'get_win_test'
]

# You could add platform checks or initial setup here if needed