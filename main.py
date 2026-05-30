"""
工资单助手  —  Wage Form Helper
Entry point: checks for required libraries, then launches the app.

Usage:
  python main.py
"""

import sys
import os

# Ensure the script's own directory is on the path so sibling modules load correctly
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Determine the app's root directory:
#   - when running as a PyInstaller exe: folder that contains the exe
#   - when running as a plain script: folder that contains main.py
if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))


def _check_deps():
    missing = []
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        missing.append("openpyxl")

    if missing:
        print("=" * 55)
        print("缺少必要的库 / Missing required libraries:")
        for lib in missing:
            print(f"    pip install {lib}")
        print("\n请在命令行运行上述命令安装后重试。")
        print("Run the command above in your terminal, then try again.")
        print("=" * 55)
        input("按 Enter 键退出… / Press Enter to exit…")
        sys.exit(1)


if __name__ == "__main__":
    _check_deps()

    import tkinter as tk
    from app import WageHelperApp

    root = tk.Tk()
    # Center window on screen
    root.update_idletasks()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    w, h = 980, 740
    root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    WageHelperApp(root, app_dir=APP_DIR)
    root.mainloop()
