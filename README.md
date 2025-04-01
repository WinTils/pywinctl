# PyWinCtl

A Python package to easily find and control Windows application windows (move, resize, minimize, maximize, close, activate, etc.).

## Features

*   Find windows by title (exact or substring match).
*   Get the active (foreground) window.
*   List all visible windows.
*   Object-oriented interface (`Window` class).
*   Control window position and size (`move_to`, `resize_to`, `move_resize`).
*   Control window state (`minimize`, `maximize`, `restore`, `close`, `activate`, `hide`, `show`).
*   Set window title (`set_title`).
*   Set always-on-top status (`set_always_on_top`).
*   Get window properties (HWND, title, position, size, visibility, active status, minimized/maximized state).
*   Get owner process ID and basic process information (requires `psutil`).
*   Robust error handling for closed/invalid windows.

## Installation

1.  **Prerequisites:**
    *   Python 3.7+
    *   Windows Operating System
    *   `pip` (Python package installer)

2.  **Install using pip:**

    ```bash
    pip install pywin32 psutil pygetwindow
    # Then install this package (if downloaded/cloned)
    pip install .
    ```
    *(Or, if published to PyPI: `pip install pywinctl`)*

## Basic Usage

```python
import pywinctl as pwc
import time
import sys

try:
    # --- Finding Windows ---

    # Find the first window with "Notepad" in the title (substring match)
    # Returns a Window object or raises WindowNotFoundError
    notepad_win = pwc.get_window_by_title("Notepad")
    print(f"Found Notepad: {notepad_win}")

    # Find a window with an exact title match
    # calc_win = pwc.get_window_by_title("Calculator", exact_match=True)
    # print(f"Found Calculator: {calc_win}") # Might fail if Calculator title isn't exact

    # Get the currently active window
    active_win = pwc.get_active_window()
    if active_win:
        print(f"Active window: {active_win.title} (HWND: {active_win.hwnd})")
    else:
        print("No active window found.")

    # List all visible windows
    all_wins = pwc.get_all_windows()
    print(f"\nFound {len(all_wins)} windows:")
    # for w in all_wins:
    #     print(f"- {w.title} ({w.size})") # Example: Print title and size

    # --- Controlling a Window (using notepad_win from above) ---
    if notepad_win:
        print(f"\nControlling '{notepad_win.title}'...")

        # Print initial state
        print(f"  Initial Pos: {notepad_win.position}, Size: {notepad_win.size}")
        print(f"  Is Active? {notepad_win.is_active}")
        print(f"  Is Minimized? {notepad_win.is_minimized}")
        print(f"  Process Info: {notepad_win.process_info}") # Requires psutil

        # Activate (bring to front)
        if not notepad_win.is_active:
            print("  Activating...")
            notepad_win.activate()
            time.sleep(1) # Give time for activation

        # Restore if minimized/maximized
        if notepad_win.is_minimized or notepad_win.is_maximized:
            print("  Restoring...")
            notepad_win.restore()
            time.sleep(0.5)

        # Move and Resize
        print("  Moving to (100, 100)...")
        notepad_win.move_to(100, 100)
        time.sleep(0.5)

        print("  Resizing to (800, 600)...")
        notepad_win.resize_to(800, 600)
        time.sleep(0.5)

        print("  Moving & Resizing to (200, 150, 600, 400)...")
        notepad_win.move_resize(200, 150, 600, 400)
        time.sleep(0.5)

        # Minimize / Maximize / Restore
        print("  Minimizing...")
        notepad_win.minimize()
        time.sleep(1)
        print(f"  Is Minimized? {notepad_win.is_minimized}")


        print("  Maximizing...")
        notepad_win.maximize()
        time.sleep(1)
        print(f"  Is Maximized? {notepad_win.is_maximized}")


        print("  Restoring...")
        notepad_win.restore()
        time.sleep(1)
        print(f"  Is Minimized? {notepad_win.is_minimized}")
        print(f"  Is Maximized? {notepad_win.is_maximized}")

        # Title and Always On Top
        original_title = notepad_win.title
        print(f"  Original title: {original_title}")
        print("  Setting title...")
        notepad_win.set_title("PyWinCtl Controlled Notepad!")
        time.sleep(1)
        print(f"  New title: {notepad_win.title}")
        notepad_win.set_title(original_title) # Restore title
        time.sleep(0.5)


        print("  Setting always on top...")
        notepad_win.set_always_on_top(True)
        time.sleep(2) # Observe effect
        print("  Disabling always on top...")
        notepad_win.set_always_on_top(False)
        time.sleep(0.5)

        # Close (use with caution!)
        # print("  Closing window...")
        # notepad_win.close()
        # time.sleep(1)
        # try:
        #     print(f"Window state after close attempt: {notepad_win.title}")
        # except pwc.InvalidWindowError:
        #     print("  Window successfully closed or invalid.")

except pwc.WindowNotFoundError:
    print("Window not found. Please ensure 'Notepad' is running.")
except pwc.PyWinCtlError as e:
    print(f"An error occurred: {e}")
except Exception as e:
     print(f"An unexpected error occurred: {e}")
     import traceback
     traceback.print_exc()