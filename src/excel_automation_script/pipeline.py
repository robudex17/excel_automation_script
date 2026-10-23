from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from .data_sheet import write_data_sheet, write_pendientes_sheet
from .destination import build_output_path, find_output_template, save_as_new_file
from .io import find_source_file, load_source
from .metadata import parse_source_filename
from .pivot import rebuild_pivots_from_reference
from .resumen import write_chart_window, write_semana_table
from .transform import filter_by_vendor, filter_pendientes, latest_delivery_month
from .weeks import semanas_in_month, validate_semana_in_month

# A progress callback receives a short, human-readable status string at
# each step -- the CLI ignores it, the GUI uses it to update its log.
ProgressCallback = Callable[[str], None]


class PipelineError(Exception):
    """A known, user-facing failure (bad source file, bad vendor, etc.),
    as opposed to an unexpected bug -- the caller shows its message
    directly, without a stack trace."""


@dataclass
class PipelineResult:
    semana: str
    date: str
    month_year: str
    source_rows: int
    vendor_name: str
    filtered_rows: int
    pendientes_rows: int
    semanas_in_month: list[int]
    output_template_name: str
    output_path: Path


def _noop(_message: str) -> None:
    pass


# -------------------------------------------------------------
# RUN: the full pipeline for one vendor, shared by the CLI and
# the GUI. Raises PipelineError on any known/expected failure;
# anything else is an unexpected bug and propagates as-is.
# -------------------------------------------------------------
def run_pipeline(vendor_name: str, progress: ProgressCallback = _noop) -> PipelineResult:
    # 1. LOAD SOURCE: find + validate + read the one source file,
    #    and parse Semana/date/month-year from its filename
    progress("Reading source file...")
    try:
        source_path = find_source_file()
        source_meta = parse_source_filename(source_path)
        df = load_source()
    except (FileNotFoundError, ValueError) as e:
        raise PipelineError(str(e)) from e

    source_semana_number = int(source_meta.semana.split()[-1])

    try:
        validate_semana_in_month(source_semana_number, source_meta.month_year)
    except ValueError as e:
        raise PipelineError(str(e)) from e

    # 2. FILTER: keep only this vendor's rows
    progress(f"Filtering rows for '{vendor_name}'...")
    filtered_df = filter_by_vendor(df, vendor_name)
    if filtered_df.empty:
        raise PipelineError(f"No rows found for vendor '{vendor_name}' in the source file.")

    # 3. LOCATE OUTPUT TEMPLATE: the single, fixed template file,
    #    then copy it under the new output name
    try:
        output_template = find_output_template()
    except FileNotFoundError as e:
        raise PipelineError(str(e)) from e

    output_path = build_output_path(vendor_name, source_meta.semana, source_meta.date)
    try:
        save_as_new_file(output_template, output_path)
    except PermissionError as e:
        raise PipelineError(
            f"Can't write to '{output_path.name}' -- it looks like the file is "
            "already open in Excel. Please close it and try again."
        ) from e

    # 4. WRITE 'Data': clear old records and write the filtered
    #    data into the 'Data' sheet of the newly copied output file
    progress("Writing Data sheet...")
    try:
        write_data_sheet(output_path, filtered_df, source_meta.date)
    except ValueError as e:
        raise PipelineError(str(e)) from e

    # 5. WRITE 'Pendientes - Contrato': further filter to active
    #    contracts still pending in the source's month, then write
    progress("Writing Pendientes - Contrato sheet...")
    pendientes_df = filter_pendientes(filtered_df, source_meta.month_year, source_meta.date)

    try:
        write_pendientes_sheet(output_path, pendientes_df, source_meta.date)
    except ValueError as e:
        raise PipelineError(str(e)) from e

    # 6. WRITE Semana table
    progress("Writing Semana table...")
    month_weeks = semanas_in_month(source_meta.month_year)

    try:
        write_semana_table(output_path, month_weeks, source_semana_number, source_meta.month_year)
    except ValueError as e:
        raise PipelineError(str(e)) from e

    # 7. WRITE chart window
    progress("Writing Resultados / chart data...")
    source_year = int(source_meta.month_year.split("-")[1])
    last_data_month = latest_delivery_month(filtered_df, source_year)

    try:
        write_chart_window(output_path, source_meta.month_year, last_data_month)
    except ValueError as e:
        raise PipelineError(str(e)) from e

    # 8. BUILD PIVOTS
    progress("Rebuilding pivot tables...")
    rebuild_pivots_from_reference(output_path)

    progress("Done.")
    return PipelineResult(
        semana=source_meta.semana,
        date=source_meta.date,
        month_year=source_meta.month_year,
        source_rows=len(df),
        vendor_name=vendor_name,
        filtered_rows=len(filtered_df),
        pendientes_rows=len(pendientes_df),
        semanas_in_month=month_weeks,
        output_template_name=output_template.name,
        output_path=output_path,
    )
