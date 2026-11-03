"""
PyInstaller entry point. Not part of the excel_automation_script package --
PyInstaller executes whatever script it's pointed at as a standalone
top-level script (no parent package), so gui.py's relative imports
(`from .io import ...`) fail if PyInstaller is pointed at it directly.
This thin wrapper imports the package properly by its full name instead.
"""

from excel_automation_script.gui import main

if __name__ == "__main__":
    main()
