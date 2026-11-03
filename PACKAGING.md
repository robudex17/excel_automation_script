# Packaging the desktop app

Builds a standalone Windows `.exe` of the GUI (`excel_automation_script.gui`)
for a laptop with no Python installed.

## Build

```
uv run pyinstaller "DIFOT Report Generator.spec" --clean --noconfirm
```

This produces `dist/DIFOT Report Generator.exe`. (`pyinstaller` is a
`build` dependency group -- `uv sync` alone won't install it; use
`uv add --group build pyinstaller` once if it's missing.)

## Assembling the folder to hand to the client

The `.exe` reads/writes its data *next to itself* (see `paths.py`), not
inside the bundled executable, so the exe alone isn't enough -- it needs
the template folder sitting beside it:

```
DIFOT Report Generator/
├── DIFOT Report Generator.exe
└── data/
    └── template/
        ├── output_template.xlsx
        └── pivot_reference.xlsx
```

Copy those two template files from this repo's `data/template/` into a
`data/template/` folder next to the `.exe`. `data/source` and
`data/destination` don't need to be created ahead of time -- the app
creates them itself on first use, and that's also where the client's
generated reports will land (the "Open output folder" button opens it).

Hand the client the whole `DIFOT Report Generator/` folder -- not just
the `.exe` on its own -- and tell her to keep the `data` folder next to
it. A desktop shortcut to the `.exe` is fine; the working directory
doesn't matter, only the exe's own location does.

## Updating the template later

If `output_template.xlsx` or `pivot_reference.xlsx` ever change, there's
no need to rebuild the `.exe` -- just replace those two files in the
client's `data/template/` folder.

## Notes

- Built and smoke-tested on Windows only; `--windowed` (no console
  window) and `sys.frozen` path handling are Windows-oriented, matching
  the client's actual machine.
- The `.spec` file is tracked in git; `build/` and `dist/` (generated
  output) are gitignored.
- The build points at `run_gui.py` at the repo root, not
  `src/excel_automation_script/gui.py` directly. PyInstaller runs
  whatever script it's given as a standalone top-level script with no
  parent package, so `gui.py`'s relative imports (`from .io import
  ...`) fail if it's used as the entry point directly. `run_gui.py`
  imports the package by its full name instead.
