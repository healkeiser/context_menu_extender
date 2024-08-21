#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Built-in
import ctypes
import os
import sys
import subprocess


def is_admin() -> bool:
    """Check if the script is running with administrative privileges.

    Returns:
        bool: True if the script is running with administrative privileges, False otherwise.
    """

    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


# Get the absolute path to `main.py`
main_script = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "main.py"
)

if is_admin():
    # If already running as admin, execute the main script
    subprocess.run([sys.executable, main_script], check=True)
else:
    # Relaunch as admin
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, main_script, None, 1
    )
