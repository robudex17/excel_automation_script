"""
Where the app's data lives, relative to the app itself rather than the
current working directory.

Under `uv run` (dev mode), that's the project root -- so data/source,
data/destination, data/template keep working exactly as before.

Packaged as a PyInstaller onefile .exe, the working directory when
double-clicked isn't reliable (it depends on how the shortcut/launcher
sets it), but sys.executable always points at wherever the .exe itself
is sitting. Data lives next to the .exe, not inside it, so the client's
generated reports are easy to find and the template is easy to update
without rebuilding.
"""

import sys
from pathlib import Path

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent.parent.parent

DATA_DIR = BASE_DIR / "data"
