# pywinctl/_exceptions.py

class PyWinCtlError(Exception):
    """Base exception for pywinctl errors."""
    pass

class WindowNotFoundError(PyWinCtlError):
    """Raised when a specific window cannot be found."""
    pass

class InvalidWindowError(PyWinCtlError):
    """Raised when trying to operate on an invalid or closed window handle."""
    pass

class WindowsAPIError(PyWinCtlError):
    """Raised when a Windows API call fails."""
    def __init__(self, message, error_code=None):
        super().__init__(message)
        self.error_code = error_code

    def __str__(self):
        if self.error_code is not None:
            return f"{super().__str__()} (Error Code: {self.error_code})"
        return super().__str__()