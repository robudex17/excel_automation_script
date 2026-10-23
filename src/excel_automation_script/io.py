import shutil
from pathlib import Path

import pandas as pd

from .paths import DATA_DIR

# -------------------------------------------------------------
# CONFIGURATION
# -------------------------------------------------------------
SOURCE_DIR = DATA_DIR / "source"
SKIP_ROWS = 4

# Headers the source file must contain; the run fails if any are missing.
EXPECTED_COLUMNS = [
    "Asset - Region (Text)",
    "Asset - Long Text",
    "Purch. Org. - Text",
    "Vendor(VN) - Key",
    "Vendor(VN) - Text",
    "IncoTerms - Key",
    "Inco term: Additional Information - Text",
    "Purchase Document Item No. - Key",
    "Purch. Group - Key",
    "Contract (Outline Agreement) No - Key",
    "Material - Key",
    "Material - Text",
    "Purchase Document Number - Key",
    "PO Item Quantity Order UOM",
    "Purch. Item Category - Text",
    "Indicator : Delivery Completion - Key",
    "PO Document Date - Key",
    "Date : Calculated Delivery - Key",
    "Date: First GR Posting - Key",
    "Date: PO Statistical Delv - Key",
    "Date : PO Change - Key",
    "No. of PO Schd Line",
    "Count of DIFOT",
    "DIFOT%",
]


# -------------------------------------------------------------
# 1. LOCATE: exactly one source .xlsx must exist in data/source
# -------------------------------------------------------------
def find_source_file() -> Path:
    candidates = [
        p for p in SOURCE_DIR.glob("*.xlsx") if not p.name.startswith("~$")
    ]
    if not candidates:
        raise FileNotFoundError(f"No .xlsx file found in {SOURCE_DIR}")
    if len(candidates) > 1:
        raise ValueError(
            f"Multiple .xlsx files found in {SOURCE_DIR}, expected exactly one: "
            f"{[p.name for p in candidates]}"
        )
    return candidates[0]


# -------------------------------------------------------------
# SET (GUI use): copy a user-picked source file into data/source,
# replacing whatever was there -- find_source_file() requires
# exactly one file, and a non-technical user picks it via a file
# dialog rather than placing it in the folder themselves.
# -------------------------------------------------------------
def set_source_file(path: Path) -> Path:
    SOURCE_DIR.mkdir(parents=True, exist_ok=True)
    for existing in SOURCE_DIR.glob("*.xlsx"):
        existing.unlink()
    destination = SOURCE_DIR / path.name
    shutil.copyfile(path, destination)
    return destination


# -------------------------------------------------------------
# 2. VALIDATE: all expected headers must be present
# -------------------------------------------------------------
def validate_columns(df: pd.DataFrame) -> None:
    missing = [c for c in EXPECTED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Source file is missing expected column(s): {missing}")


# -------------------------------------------------------------
# 3. LOAD: read the source file, skipping the report's title rows
# -------------------------------------------------------------
def load_source() -> pd.DataFrame:
    path = find_source_file()
    df = pd.read_excel(path, skiprows=SKIP_ROWS)
    validate_columns(df)
    return df
