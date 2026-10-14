"""Small Windows helpers: console attach/alloc for the windowed exe, message boxes, start-with-Windows."""
import ctypes
import os
import sys

RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
RUN_NAME = "SharePointBridge"


def has_console():
    """True when print()/input() reach a real console (not the windowed exe's log-only stand-in)."""
    return sys.stdout is not None and sys.stdin is not None and not getattr(sys.stdout, "is_log_writer", False)


def ensure_console():
    """The windowed exe has no console. Attach to the launching terminal or open a new window for prompts."""
    if os.name != "nt" or has_console():
        return
    k32 = ctypes.windll.kernel32
    if not k32.AttachConsole(-1):  # -1 = parent process's console
        k32.AllocConsole()
    sys.stdout = sys.stderr = open("CONOUT$", "w", buffering=1, encoding="utf-8")
    sys.stdin = open("CONIN$", "r", encoding="utf-8")


def release_console():
    """Close the temporary console window opened for setup."""
    if os.name == "nt":
        ctypes.windll.kernel32.FreeConsole()


def message_box(title, text, error=False):
    if os.name == "nt":
        ctypes.windll.user32.MessageBoxW(0, text, title, 0x10 if error else 0x40)
    else:
        print(f"{title}: {text}")


def autostart_enabled():
    if os.name != "nt":
        return False
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as key:
            winreg.QueryValueEx(key, RUN_NAME)
            return True
    except OSError:
        return False


def set_autostart(enabled):
    """Start the bridge automatically when you sign in to Windows (packaged exe only)."""
    import winreg
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as key:
        if enabled:
            winreg.SetValueEx(key, RUN_NAME, 0, winreg.REG_SZ, f'"{sys.executable}"')
        else:
            try:
                winreg.DeleteValue(key, RUN_NAME)
            except OSError:
                pass


def copy_to_clipboard(text):
    import subprocess
    subprocess.run(["clip"], input=text.encode("utf-8"), creationflags=0x08000000)  # CREATE_NO_WINDOW
