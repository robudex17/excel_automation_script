import re
import shutil
from pathlib import Path

from .paths import DATA_DIR

# -------------------------------------------------------------
# CONFIGURATION
# -------------------------------------------------------------
DESTINATION_DIR = DATA_DIR / "destination"
TEMPLATE_PATH = DATA_DIR / "template" / "output_template.xlsx"


# -------------------------------------------------------------
# HELPER: strip characters that are invalid in Windows filenames
# -------------------------------------------------------------
def _sanitize(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "", name).strip()


# -------------------------------------------------------------
# 1. LOCATE: the single, fixed output template (client-provided,
#    not searched by vendor/Semana)
# -------------------------------------------------------------
def find_output_template() -> Path:
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Output template not found: {TEMPLATE_PATH}")
    return TEMPLATE_PATH


# -------------------------------------------------------------
# 2. BUILD OUTPUT NAME: "<vendor> - <Semana N> - <DDMonYYYY>.xlsx"
#    Reuses the source file's own Semana and raw date, e.g.
#    "SSAB SWEDISH STEEL ACEROS DE CHILE - Semana 38 - 14Sep2026.xlsx"
# -------------------------------------------------------------
def build_output_path(vendor_name: str, source_semana_label: str, source_date: str) -> Path:
    clean_vendor = _sanitize(vendor_name)
    filename = f"{clean_vendor} - {source_semana_label} - {source_date}.xlsx"
    return DESTINATION_DIR / filename


# -------------------------------------------------------------
# 3. SAVE: copy the template to the new file, leaving the
#    original output template untouched
# -------------------------------------------------------------
def save_as_new_file(template: Path, output_path: Path) -> None:
    shutil.copy(template, output_path)
