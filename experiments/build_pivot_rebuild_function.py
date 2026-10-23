"""
EXPERIMENT round 3 (isolated, disposable): the real function that will
eventually move into src/excel_automation_script/pivot.py once verified.

Idea: don't trust the pivot baked into the live output_template.xlsx to stay
intact (it's degraded twice this session just from Excel-side interaction).
Instead, every run, REPLACE whatever pivot is in the output file's
'Resumen - Contrato' / 'Resumen - Todo' sheets with a fresh clone from a
dedicated, never-opened-in-Excel reference file
(data/template/pivot_reference.xlsx). Self-healing: even if the live
template's own pivot degrades, every generated output still gets a
known-good pivot injected.

Key fix vs. round 2: openpyxl loads each pivot's cache as a SEPARATE Python
object even when they share one cacheId on disk (verified: p1.cache is not
p2.cache for the two real pivots, despite both being cacheId=59). So the two
clones must be made to explicitly share ONE cache object -- otherwise we
hit the same duplicate-cacheId collision that broke round 1.

Does NOT touch src/excel_automation_script or any real data/ file except
reading data/template/pivot_reference.xlsx (never written to).
"""

import copy
from pathlib import Path
import shutil

import openpyxl

REFERENCE_FILE = Path("data/template/pivot_reference.xlsx")
TARGET_SOURCE = Path("data/destination/SSAB SWEDISH STEEL ACEROS DE CHILE - Semana 27 - 30Jun2026.xlsx")
OUTPUT_FILE = Path("experiments/pivot_rebuild_test.xlsx")


def rebuild_pivots_from_reference(workbook_path: Path, reference_path: Path) -> None:
    ref_wb = openpyxl.load_workbook(reference_path)
    wb = openpyxl.load_workbook(workbook_path)

    ref_items = [
        (sheet_name, pivot)
        for sheet_name in ref_wb.sheetnames
        for pivot in getattr(ref_wb[sheet_name], "_pivots", [])
    ]
    if not ref_items:
        raise ValueError(f"No pivot tables found in reference file: {reference_path}")

    shared_cache = copy.deepcopy(ref_items[0][1].cache)
    shared_cache.refreshOnLoad = True

    for sheet_name, ref_pivot in ref_items:
        if sheet_name not in wb.sheetnames:
            continue
        target_ws = wb[sheet_name]
        target_ws._pivots = []

        cloned = copy.deepcopy(ref_pivot)
        cloned.cacheId = 1
        cloned.cache = shared_cache
        for field in cloned.pivotFields:
            if field.axis == "axisPage":
                field.includeNewItemsInFilter = True
                field.hideNewItems = False

        target_ws.add_pivot(cloned)

    wb.save(workbook_path)


def main() -> None:
    if not REFERENCE_FILE.exists():
        raise FileNotFoundError(f"Reference not found: {REFERENCE_FILE}")
    if not TARGET_SOURCE.exists():
        raise FileNotFoundError(f"Target source not found: {TARGET_SOURCE}")

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(TARGET_SOURCE, OUTPUT_FILE)

    rebuild_pivots_from_reference(OUTPUT_FILE, REFERENCE_FILE)
    print(f"Saved: {OUTPUT_FILE}")

    # -----------------------------------------------------------------
    # Verify: reload and check structure, especially cacheId sharing
    # -----------------------------------------------------------------
    wb2 = openpyxl.load_workbook(OUTPUT_FILE)
    p1 = wb2["Resumen - Contrato"]._pivots[0]
    p2 = wb2["Resumen - Todo"]._pivots[0]
    print(f"Resumen - Contrato: name={p1.name} location={p1.location.ref} cacheId={p1.cacheId}")
    print(f"Resumen - Todo:     name={p2.name} location={p2.location.ref} cacheId={p2.cacheId}")
    print(f"rowFields: {[f.x for f in p1.rowFields]}  colFields: {[f.x for f in p1.colFields]}")
    page_axis_fields = [f for f in p1.pivotFields if f.axis == "axisPage"]
    print(f"pageFields: {[f.fld for f in p1.pageFields]}")
    print(f"page-axis pivotField includeNewItemsInFilter: {[f.includeNewItemsInFilter for f in page_axis_fields]}")


if __name__ == "__main__":
    main()
