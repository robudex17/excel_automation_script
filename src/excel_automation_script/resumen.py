import re
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter

TOP_BORDER = Border(top=Side(style="thin", color="000000"))
NO_FILL = PatternFill(fill_type=None)
# Highlight fills, matching the template's original theme colors/tint
# (same theme/tint as the 'Resultados' header at B18).
HEADER_FILL = PatternFill(patternType="solid", fgColor=Color(theme=4, tint=0.7999816888943144))
HEADER_FONT = Font(bold=True)
# Light green, marking the whole row for the Semana currently being generated.
CURRENT_SEMANA_ROW_FILL = PatternFill(patternType="solid", fgColor="C6EFCE")
TOTAL_PROV_FILL = PatternFill(patternType="solid", fgColor=Color(theme=7, tint=0.7999816888943144))
PERCENT_FORMAT = "0.0%"
CENTER_ALIGN = Alignment(horizontal="center")

# -------------------------------------------------------------
# CONFIGURATION
# -------------------------------------------------------------
RESUMEN_CONTRATO_SHEET = "Resumen - Contrato"
RESUMEN_TODO_SHEET = "Resumen - Todo"
TABLE_START_ROW = 3
TABLE_HEADERS = ["Semana", "PO Pendiente", "Compromiso lineas", "Compromiso prov. %", "Oportunidad Real"]
# Fixed row of the 'OK' line in the Resultados/META/OK/NOK/EN PROCESO/OPORTUNIDAD
# block below the pivot (rows 18-23 in the template); only its column varies,
# by month.
OK_ROW = 20
# Fixed starting column for the Semana table -- column R, matching where the
# client's own manually-built reference file places it. NOT derived from the
# pivot's cached bounds: that cache only reflects whatever state Excel last
# saved it in (which vendor, or even a near-empty state after repeated
# edits during development), so it can't be trusted to compute a safe
# position from. A fixed column, sized generously past the pivot's maximum
# realistic width (Jan-Dec + Total), is simple and never guesses wrong.
SEMANA_TABLE_START_COL = 18  # R
# Start of the safety zone to clear stale content from (values, borders,
# fills left behind by a table at some earlier, different position during
# development) -- covers well before and after SEMANA_TABLE_START_COL.
CLEAR_ZONE_START_COL = 16  # P
CLEAR_ZONE_COLUMNS = 30

# The 'Data' sheet's structured Table and the columns these formulas read
# from -- computed directly against the Data table (SUMIFS), not against
# the pivot's cached cells. The pivot cache only reflects whatever vendor
# was last actually refreshed by a human in Excel, and its row layout
# (which DELIVERY sub-categories even appear) varies by vendor, so
# hardcoded pivot-cell/row-offset references silently point at the wrong
# numbers for any vendor other than the one the cache happens to be
# frozen on. SUMIFS against the table is correct regardless.
DATA_TABLE = "Table1"
COL_PO_SCHD_LINE = "No. of PO Schd Line"
COL_MES = "MES"
COL_CONTRACT = "Contract (Outline Agreement) No - Key"
COL_ANO = "AÑO"
COL_DELIVERY = "DELIVERY"
COL_PERFORMANCE = "PERFORMANCE"

MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


# -------------------------------------------------------------
# HELPER: a loaded cell's border never equals a bare Border()
# (openpyxl's shared-style representation), so check the actual
# side styles instead of comparing the whole object
# -------------------------------------------------------------
def _has_no_border(border: Border) -> bool:
    return all(
        getattr(border, side).style is None
        for side in ("left", "right", "top", "bottom")
    )


# -------------------------------------------------------------
# HELPER: SUMIFS of No. of PO Schd Line against Table1, filtered
# by MES (month-year, e.g. 'Aug-2026') and optionally DELIVERY/
# PERFORMANCE. Wrapped in IFERROR so a month with zero rows
# (nothing yet, or nothing ever) safely divides to 0 rather than
# #DIV/0!.
# -------------------------------------------------------------
def _sumifs(
    month_year: str | None = None,
    performance: str | None = None,
    delivery: str | None = None,
    year: str | None = None,
    exclude_no_contract: bool = True,
) -> str:
    # Contrato excludes rows with no contract ('#'), matching the client's
    # pivot (its Contract page filter hides '#') and Pendientes; Todo keeps them.
    criteria = []
    if exclude_no_contract:
        criteria.append(f'{DATA_TABLE}[{COL_CONTRACT}],"<>#"')
    if month_year is not None:
        criteria.append(f'{DATA_TABLE}[{COL_MES}],"{month_year}"')
    if year is not None:
        criteria.append(f'{DATA_TABLE}[{COL_ANO}],"{year}"')
    if performance is not None:
        criteria.append(f'{DATA_TABLE}[{COL_PERFORMANCE}],"{performance}"')
    if delivery is not None:
        criteria.append(f'{DATA_TABLE}[{COL_DELIVERY}],"{delivery}"')
    return f"SUMIFS({DATA_TABLE}[{COL_PO_SCHD_LINE}],{','.join(criteria)})"


# -------------------------------------------------------------
# WRITE: (re)build the Semana/PO Pendiente/Compromiso/Oportunidad
# table on 'Resumen - Contrato', at a fixed position (SEMANA_TABLE_
# START_COL) matching the client's own manual layout -- not derived
# from the pivot's cache, which can't be trusted to reflect a safe,
# current position (see the constant's comment above). Shows every
# Semana that belongs to the source's month (4 or 5 weeks, e.g.
# Sep-2026 -> SEM 36..40), highlighting whichever row matches the
# source's actual current Semana (not necessarily the list's last
# entry -- a month can include weeks later than the report's "as
# of" week). 'Compromiso prov. %' and 'Oportunidad Real' are
# computed via SUMIFS against the Data table for the source's
# month, and by referencing the chart block's own OK row at that
# month's fixed column (MONTH_ORDER-based) -- not the pivot's cache.
# -------------------------------------------------------------
def write_semana_table(workbook_path: Path, weeks: list[int], current_semana: int, month_year: str) -> None:
    wb = openpyxl.load_workbook(workbook_path)

    if RESUMEN_CONTRATO_SHEET not in wb.sheetnames:
        raise ValueError(f"'{RESUMEN_CONTRATO_SHEET}' sheet not found in {workbook_path.name}")

    ws = wb[RESUMEN_CONTRATO_SHEET]
    start_col = SEMANA_TABLE_START_COL

    month_abbrev = month_year.split("-")[0]
    ok_row_col_letter = get_column_letter(CHART_WINDOW_START_COL + MONTH_ORDER.index(month_abbrev))
    month_total_expr = _sumifs(month_year)

    # 1. Clear a generous zone around the fixed position, wiping any
    #    leftover table (values, borders, and fills) from a previous
    #    run's (possibly different, pre-fix) position or a different-length
    #    week list. Only touch cells that actually have something to clear,
    #    so pristine cells aren't dragged into the sheet's used range.
    for row in range(TABLE_START_ROW, TABLE_START_ROW + len(weeks) + 10):
        for col in range(CLEAR_ZONE_START_COL, CLEAR_ZONE_START_COL + CLEAR_ZONE_COLUMNS):
            cell = ws.cell(row=row, column=col)
            if cell.value is None and cell.fill.fill_type is None and _has_no_border(cell.border):
                continue
            cell.value = None
            cell.border = Border()
            cell.fill = NO_FILL

    label_col, po_col, lineas_col, prov_col, oport_col = range(start_col, start_col + 5)
    label_letter = get_column_letter(label_col)
    po_letter = get_column_letter(po_col)
    lineas_letter = get_column_letter(lineas_col)
    prov_letter = get_column_letter(prov_col)
    oport_letter = get_column_letter(oport_col)

    # 2. Headers
    for offset, header in enumerate(TABLE_HEADERS):
        cell = ws.cell(row=TABLE_START_ROW, column=start_col + offset, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT

    # 3. Week labels + formulas, one row per Semana in the source's month
    first_row = TABLE_START_ROW + 1
    last_row = first_row + len(weeks) - 1
    current_semana_row = None
    for i, week in enumerate(weeks):
        row = first_row + i
        if week == current_semana:
            current_semana_row = row
        ws.cell(row=row, column=label_col, value=f"SEM {week}")

        ws.cell(row=row, column=po_col, value=(
            f"=COUNTIF('Pendientes - Contrato'!$Y:AE,{label_letter}{row})"
        ))
        ws.cell(row=row, column=lineas_col, value=(
            f"={po_letter}{row}-(COUNTIFS('Pendientes - Contrato'!$Y:AD,{label_letter}{row},"
            f"'Pendientes - Contrato'!$AE:AJ,\"no\"))"
        ))
        prov_cell = ws.cell(row=row, column=prov_col, value=f"=IFERROR({lineas_letter}{row}/{month_total_expr},0)")
        prov_cell.number_format = PERCENT_FORMAT
        if row == first_row:
            oport_cell = ws.cell(
                row=row, column=oport_col,
                value=f"=${ok_row_col_letter}${OK_ROW}+{prov_letter}{row}",
            )
        else:
            oport_cell = ws.cell(row=row, column=oport_col, value=f"={oport_letter}{row - 1}+{prov_letter}{row}")
        oport_cell.number_format = PERCENT_FORMAT

    if current_semana_row is None:
        raise ValueError(
            f"Semana {current_semana} not found among this month's weeks: {weeks}"
        )

    # 4. Total row, set off from the week rows above it with a top border
    total_row = last_row + 1
    ws.cell(row=total_row, column=label_col, value="Total")
    ws.cell(row=total_row, column=po_col, value=f"=SUM({po_letter}{first_row}:{po_letter}{last_row})")
    ws.cell(row=total_row, column=lineas_col, value=f"=SUM({lineas_letter}{first_row}:{lineas_letter}{last_row})")
    total_prov_cell = ws.cell(row=total_row, column=prov_col, value=f"=SUM({prov_letter}{first_row}:{prov_letter}{last_row})")
    total_prov_cell.number_format = PERCENT_FORMAT
    # Oportunidad Real is a running cumulative total, so its "Total" is just
    # the last week's value carried down (summing every row would double-count).
    total_oport_cell = ws.cell(row=total_row, column=oport_col, value=f"={oport_letter}{last_row}")
    total_oport_cell.number_format = PERCENT_FORMAT

    for col in range(label_col, oport_col + 1):
        cell = ws.cell(row=total_row, column=col)
        cell.border = TOP_BORDER
        cell.alignment = CENTER_ALIGN

    # 5. Highlights: the entire row for the Semana currently being generated,
    #    and Total's Compromiso prov. % / Oportunidad Real
    for col in range(label_col, oport_col + 1):
        ws.cell(row=current_semana_row, column=col).fill = CURRENT_SEMANA_ROW_FILL
    total_prov_cell.fill = TOTAL_PROV_FILL
    total_oport_cell.fill = TOTAL_PROV_FILL

    wb.save(workbook_path)


# -------------------------------------------------------------
# CONFIGURATION -- the chart-feeding block below the pivot
# -------------------------------------------------------------
CHART_WINDOW_START_COL = 4  # D -- the chart's own cat/val range always starts here
CHART_WINDOW_MAX_SIZE = 12  # D:O -- the absolute cap (Jan-Dec); the chart's actual
                            # range ends exactly at the current month, growing one
                            # column per month rather than always spanning the full year.
CHART_LABEL_ROW = 18        # Resultados (month labels, feeds the chart's category axis)
CHART_META_ROW = 19         # META (constant target line, e.g. 0.94)
CHART_OK_ROW = 20           # OK
CHART_NOK_ROW = 21          # NOK
CHART_EN_PROCESO_ROW = 22   # EN PROCESO
CHART_OPORTUNIDAD_ROW = 23  # OPORTUNIDAD (feeds the chart)
# How many columns past CHART_WINDOW_START_COL to clear stale content from,
# wider than CHART_WINDOW_MAX_SIZE (the real Jan-Dec cap) so leftover
# formulas from an earlier run/code version -- e.g. a stray column just
# past Dec -- always get wiped, not just the 12 columns we actually write to.
CHART_CLEAR_ZONE_SIZE = 20


# -------------------------------------------------------------
# HELPER: widen an existing plain cell-range reference's end
# column, keeping its row and start column untouched, e.g.
# "'Sheet'!$D$19:$K$19" -> "'Sheet'!$D$19:$O$19"
# -------------------------------------------------------------
def _widen_range_end_column(ref: str, new_end_col_letter: str) -> str:
    start, end = ref.rsplit(":", 1)
    row_num = re.search(r"\d+$", end).group()
    return f"{start}:${new_end_col_letter}${row_num}"


# -------------------------------------------------------------
# WRITE: populate the chart-feeding table (rows 18-23) for all
# 12 months (Jan-Dec), computed via SUMIFS directly against the
# Data table for each month -- since PO delivery dates already
# extend months into the future, October-December can show real
# (or correctly zero) values even while the report is only "as
# of" September. The chart's own cat/val ranges are resized to match
# the table: whichever is later, the current reporting month or this
# vendor's last scheduled delivery month -- by extending or shrinking
# their existing plain cell reference's end column -- not switched to
# a named range/OFFSET, which is what broke
# rendering before, and not read from the pivot's cache, which
# is what produced wrong numbers for any vendor other than
# whichever one the cache happened to be frozen on. The table
# itself is capped at whichever is later: the current reporting
# month, or the latest month this vendor actually has any
# delivery scheduled for (last_data_month) -- matching how far
# the pivot table itself naturally extends, since some vendors'
# real PO commitments don't reach as far into the year as others.
# -------------------------------------------------------------
def _write_chart_block(
    ws, month_year: str, last_data_month: int, exclude_no_contract: bool, total_col: int | None,
) -> None:
    current_month_abbrev, year = month_year.split("-")
    chart_months_count = MONTH_ORDER.index(current_month_abbrev) + 1  # Jan..current month
    table_months_count = min(max(chart_months_count, last_data_month), CHART_WINDOW_MAX_SIZE)

    for i in range(CHART_CLEAR_ZONE_SIZE):
        target_col = CHART_WINDOW_START_COL + i

        if i >= table_months_count:
            for row in (
                CHART_LABEL_ROW, CHART_META_ROW, CHART_OK_ROW,
                CHART_NOK_ROW, CHART_EN_PROCESO_ROW, CHART_OPORTUNIDAD_ROW,
            ):
                ws.cell(row=row, column=target_col).value = None
            continue

        month_name = MONTH_ORDER[i]
        this_month_year = f"{month_name}-{year}"

        total_expr = _sumifs(this_month_year, exclude_no_contract=exclude_no_contract)
        ok_expr = _sumifs(this_month_year, performance="OK", delivery="COMPLETA", exclude_no_contract=exclude_no_contract)
        # NOK covers every delivery status (COMPLETA, PARCIAL, PENDIENTE), as the client's NOK Total does
        nok_expr = _sumifs(this_month_year, performance="NOK", exclude_no_contract=exclude_no_contract)
        ok_parcial_expr = _sumifs(this_month_year, performance="OK", delivery="PARCIAL", exclude_no_contract=exclude_no_contract)
        ok_pendiente_expr = _sumifs(this_month_year, performance="OK", delivery="PENDIENTE", exclude_no_contract=exclude_no_contract)

        ws.cell(row=CHART_LABEL_ROW, column=target_col, value=month_name)
        ws.cell(row=CHART_META_ROW, column=target_col, value=0.94)
        ws.cell(row=CHART_OK_ROW, column=target_col, value=f"=IFERROR({ok_expr}/{total_expr},0)")
        ws.cell(row=CHART_NOK_ROW, column=target_col, value=f"=IFERROR({nok_expr}/{total_expr},0)")
        ws.cell(
            row=CHART_EN_PROCESO_ROW, column=target_col,
            value=f"=IFERROR(({ok_parcial_expr}+{ok_pendiente_expr})/{total_expr},0)",
        )
        tgt = get_column_letter(target_col)
        ws.cell(row=CHART_OPORTUNIDAD_ROW, column=target_col, value=f"={tgt}{CHART_OK_ROW}+{tgt}{CHART_EN_PROCESO_ROW}")

    # Total column: the whole-year ratios, computed directly from Table1 for
    # the source's year. Same meaning as the client's P-column formulas
    # (P20=P12/P16, P21=P11/P16, P22=(P13+P14)/P16, P23=P20+P22) but without
    # depending on pivot cell positions, which move with each vendor's months.
    if total_col is not None:
        tgt = get_column_letter(total_col)
        year_total = _sumifs(year=year, exclude_no_contract=exclude_no_contract)

        def ratio(**criteria) -> str:
            return f"=IFERROR({_sumifs(year=year, exclude_no_contract=exclude_no_contract, **criteria)}/{year_total},0)"

        ws.cell(row=CHART_LABEL_ROW, column=total_col, value="Total")
        ws.cell(row=CHART_META_ROW, column=total_col, value=0.94)
        ws.cell(row=CHART_OK_ROW, column=total_col, value=ratio(performance="OK", delivery="COMPLETA"))
        ws.cell(row=CHART_NOK_ROW, column=total_col, value=ratio(performance="NOK"))
        ws.cell(
            row=CHART_EN_PROCESO_ROW, column=total_col,
            value=(
                f"=IFERROR(({_sumifs(year=year, performance='OK', delivery='PARCIAL', exclude_no_contract=exclude_no_contract)}"
                f"+{_sumifs(year=year, performance='OK', delivery='PENDIENTE', exclude_no_contract=exclude_no_contract)})/{year_total},0)"
            ),
        )
        ws.cell(row=CHART_OPORTUNIDAD_ROW, column=total_col, value=f"={tgt}{CHART_OK_ROW}+{tgt}{CHART_EN_PROCESO_ROW}")
        # Match the month cells' styling (percent format, centered)
        for row in (CHART_OK_ROW, CHART_NOK_ROW, CHART_EN_PROCESO_ROW, CHART_OPORTUNIDAD_ROW):
            ws.cell(row=row, column=total_col)._style = copy(ws.cell(row=row, column=total_col - 1)._style)
        ws.cell(row=CHART_LABEL_ROW, column=total_col)._style = copy(ws.cell(row=CHART_LABEL_ROW, column=total_col - 1)._style)

    # Resize the chart's own cat/val ranges to match the table: whichever is
    # later, the current reporting month or this vendor's last scheduled
    # delivery month (table_months_count) -- extends or shrinks the existing
    # plain range's end column, leaving its type (a direct cell reference,
    # not a named range) untouched.
    window_end_col_letter = get_column_letter(CHART_WINDOW_START_COL + table_months_count - 1)
    for chart in ws._charts:
        for series in chart.series:
            cat_ref = series.cat.strRef if series.cat.strRef is not None else series.cat.numRef
            if cat_ref is not None:
                cat_ref.f = _widen_range_end_column(cat_ref.f, window_end_col_letter)
            if series.val.numRef is not None:
                series.val.numRef.f = _widen_range_end_column(series.val.numRef.f, window_end_col_letter)


# -------------------------------------------------------------
# CONFIGURATION -- which Resumen sheets get the chart block:
#   Contrato: excludes contract '#' rows (matches the client's pivot)
#   Todo: keeps every row, and has a Total column (P) at the end
# -------------------------------------------------------------
RESULTADOS_SHEETS = [
    {"sheet": RESUMEN_CONTRATO_SHEET, "exclude_no_contract": True, "total_col": 16},  # P
    {"sheet": RESUMEN_TODO_SHEET, "exclude_no_contract": False, "total_col": 16},  # P
]


def write_chart_window(workbook_path: Path, month_year: str, last_data_month: int) -> None:
    wb = openpyxl.load_workbook(workbook_path)

    for config in RESULTADOS_SHEETS:
        if config["sheet"] not in wb.sheetnames:
            raise ValueError(f"'{config['sheet']}' sheet not found in {workbook_path.name}")
        _write_chart_block(
            wb[config["sheet"]], month_year, last_data_month,
            exclude_no_contract=config["exclude_no_contract"],
            total_col=config["total_col"],
        )

    wb.save(workbook_path)
