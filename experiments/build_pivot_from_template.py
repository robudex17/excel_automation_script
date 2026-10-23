"""
EXPERIMENT round 2 (isolated, disposable): the hand-built-from-scratch pivot
(build_pivot_from_scratch.py) threw an error on open in real Excel -- our
manually stubbed-out PivotField.items / empty CacheField.sharedItems were
almost certainly schema-invalid (openpyxl's reader is lenient; Excel's isn't).

New approach: instead of authoring pivot XML by hand, CLONE the real,
Excel-authored pivot that already lives in data/template/output_template.xlsx
('Resumen - Contrato' sheet) -- which, as it happens, is *already* configured
exactly to the user's spec (rowFields=[PERFORMANCE, DELIVERY],
colFields=[Years, Months], pageFields=[Contract], dataFields=[Sum of No. of
PO Schd Line]) -- and just relocate the copy onto a new sheet with a fresh
cacheId. Every sharedItems/groupItems/item list carried over is genuine,
Excel-produced data, not hand-typed stubs. Only rowItems/colItems (the
computed aggregation results) are cleared, relying on refreshOnLoad to
recompute them for whatever Data ends up in this copy.

Does NOT touch src/excel_automation_script, data/template/output_template.xlsx,
or any other real file -- everything happens on a throwaway copy under
experiments/. Delete this whole experiments/ folder any time to roll back.
"""

import copy
from pathlib import Path
import shutil

import openpyxl

TEMPLATE_FILE = Path("data/template/output_template.xlsx")
SOURCE_DATA_FILE = Path("data/destination/CML COMERCIALIZADORA E IMPORTADORA - Semana 27 - 30Jun2026.xlsx")
OUTPUT_FILE = Path("experiments/pivot_test_from_template.xlsx")
PIVOT_SHEET_NAME = "Pivot Test"
DATA_SHEET_NAME = "Data"
DATA_TABLE_NAME = "Table1"


def copy_data_sheet(dst_wb: openpyxl.Workbook, src_path: Path) -> None:
    """Copy Data sheet cell content (values + formulas) from a real
    generated output file into dst_wb's Data sheet, then resize Table1."""
    src_wb = openpyxl.load_workbook(src_path)
    src_ws = src_wb[DATA_SHEET_NAME]
    dst_ws = dst_wb[DATA_SHEET_NAME]

    if dst_ws.max_row > 1:
        dst_ws.delete_rows(2, dst_ws.max_row)

    for row in src_ws.iter_rows(min_row=2, max_row=src_ws.max_row):
        for cell in row:
            dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)

    last_col_letter = dst_ws.cell(row=1, column=dst_ws.max_column).column_letter
    dst_ws.tables[DATA_TABLE_NAME].ref = f"A1:{last_col_letter}{src_ws.max_row}"


def main() -> None:
    if not TEMPLATE_FILE.exists():
        raise FileNotFoundError(f"Template not found: {TEMPLATE_FILE}")
    if not SOURCE_DATA_FILE.exists():
        raise FileNotFoundError(f"Source data file not found: {SOURCE_DATA_FILE}")

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(TEMPLATE_FILE, OUTPUT_FILE)

    wb = openpyxl.load_workbook(OUTPUT_FILE)

    copy_data_sheet(wb, SOURCE_DATA_FILE)

    real_pivot = wb["Resumen - Contrato"]._pivots[0]
    print(f"Cloning real pivot '{real_pivot.name}' (location {real_pivot.location.ref})")

    cloned = copy.deepcopy(real_pivot)
    cloned.name = "PivotFromTemplateClone"
    # The clone got a *new* CacheDefinition object (via deepcopy) but kept
    # the OLD cacheId attribute value, which the real pivots still use for
    # their (shared, unchanged) cache. Two different caches both claiming
    # the same cacheId in workbook.xml is a real collision -- almost
    # certainly what made round 1's file error out. Give the clone its own.
    cloned.cacheId = real_pivot.cacheId + 1
    cloned.location.ref = "B3:P14"
    cloned.location.firstHeaderRow = 1
    cloned.location.firstDataRow = 3
    cloned.location.firstDataCol = 2
    # Clear the computed aggregation hierarchy -- it belonged to the old
    # location/vendor snapshot. refreshOnLoad=True asks Excel to recompute
    # it fresh, same as it already does for the real pivots on every run.
    cloned.rowItems = []
    cloned.colItems = []
    cloned.cache.refreshOnLoad = True

    if PIVOT_SHEET_NAME in wb.sheetnames:
        del wb[PIVOT_SHEET_NAME]
    pivot_ws = wb.create_sheet(PIVOT_SHEET_NAME)
    pivot_ws.add_pivot(cloned)

    wb.save(OUTPUT_FILE)
    print(f"Saved: {OUTPUT_FILE}")

    # -----------------------------------------------------------------
    # Verify: reload and check the structure survives a round trip
    # -----------------------------------------------------------------
    wb2 = openpyxl.load_workbook(OUTPUT_FILE)
    reloaded = wb2[PIVOT_SHEET_NAME]._pivots
    print(f"Pivots on reload: {len(reloaded)}")
    if reloaded:
        p = reloaded[0]
        print(f"  name: {p.name}")
        print(f"  location.ref: {p.location.ref}")
        print(f"  rowFields: {[f.x for f in p.rowFields]}")
        print(f"  colFields: {[f.x for f in p.colFields]}")
        print(f"  pageFields: {[(f.fld, f.item) for f in p.pageFields]}")
        print(f"  dataFields: {[(f.name, f.fld, f.subtotal) for f in p.dataFields]}")
        print(f"  cache.refreshOnLoad: {p.cache.refreshOnLoad}")
        print(f"  cache.cacheFields count: {len(p.cache.cacheFields)}")
    print(f"Data sheet rows: {wb2[DATA_SHEET_NAME].max_row}")
    print(f"Original 'Resumen - Contrato' pivot untouched: {wb2['Resumen - Contrato']._pivots[0].location.ref}")


if __name__ == "__main__":
    main()
