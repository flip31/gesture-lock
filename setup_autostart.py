"""
setup_autostart.py
──────────────────
Run this ONCE to make GestureLock start automatically every time Windows boots.
It adds an entry to the Windows registry under:
  HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run

Usage:
  python setup_autostart.py          ← enable auto-start
  python setup_autostart.py --remove ← disable auto-start
"""

import sys
import os
import winreg

APP_NAME  = "GestureLock"
# Points to pythonw.exe (no console window) running gesture_lock.py
SCRIPT    = os.path.abspath(os.path.join(os.path.dirname(__file__), "gesture_lock.py"))
PYTHONW   = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
REG_PATH  = r"Software\Microsoft\Windows\CurrentVersion\Run"
CMD       = f'"{PYTHONW}" "{SCRIPT}"'


def enable():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH, 0, winreg.KEY_SET_VALUE)
    winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, CMD)
    winreg.CloseKey(key)
    print(f"✅ GestureLock will now start automatically on Windows login.")
    print(f"   Command registered: {CMD}")


def disable():
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH, 0, winreg.KEY_SET_VALUE)
        winreg.DeleteValue(key, APP_NAME)
        winreg.CloseKey(key)
        print("✅ GestureLock auto-start removed.")
    except FileNotFoundError:
        print("ℹ️  GestureLock was not set to auto-start.")


if __name__ == "__main__":
    if "--remove" in sys.argv:
        disable()
    else:
        enable()