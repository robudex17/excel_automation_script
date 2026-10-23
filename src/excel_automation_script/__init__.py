import argparse
import sys

from .data_sheet import write_data_sheet, write_pendientes_sheet
from .destination import build_output_path, find_output_template, save_as_new_file
from .io import find_source_file, load_source
from .metadata import parse_source_filename
from .pivot import rebuild_pivots_from_reference
from .resumen import write_chart_window, write_semana_table
from .transform import filter_by_vendor, filter_pendientes, latest_delivery_month
from .weeks import semanas_in_month, validate_semana_in_month


def main() -> None:
    # -------------------------------------------------------------
    # 1. CLI ARGS
    # -------------------------------------------------------------
    parser = argparse.ArgumentParser(description="Excel automation script")
    parser.add_argument("vendor_name", help="Vendor name to filter the source data by")
    args = parser.parse_args()

    # -------------------------------------------------------------
    # 2. LOAD SOURCE: find + validate + read the one source file,
    #    and parse Semana/date/month-year from its filename
    # -------------------------------------------------------------
    try:
        source_path = find_source_file()
        source_meta = parse_source_filename(source_path)
        df = load_source()
    except (FileNotFoundError, ValueError) as e:
        print(f"Error: {e}")
        sys.exit(1)

    source_semana_number = int(source_meta.semana.split()[-1])

    try:
        validate_semana_in_month(source_semana_number, source_meta.month_year)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)

    # -------------------------------------------------------------
    # 3. FILTER: keep only this vendor's rows
    # -------------------------------------------------------------
    filtered_df = filter_by_vendor(df, args.vendor_name)

    # -------------------------------------------------------------
    # 4. LOCATE OUTPUT TEMPLATE: the single, fixed template file,
    #    then copy it under the new output name
    # -------------------------------------------------------------
    try:
        output_template = find_output_template()
    except FileNotFoundError as e:
        print(f"Error: {e}")
        sys.exit(1)

    output_path = build_output_path(args.vendor_name, source_meta.semana, source_meta.date)
    save_as_new_file(output_template, output_path)

    # -------------------------------------------------------------
    # 5. WRITE 'Data': clear old records and write the filtered
    #    data into the 'Data' sheet of the newly copied output file
    # -------------------------------------------------------------
    try:
        write_data_sheet(output_path, filtered_df, source_meta.date)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)

    # -------------------------------------------------------------
    # 6. WRITE 'Pendientes - Contrato': further filter to active
    #    contracts still pending in the source's month, then write
    # -------------------------------------------------------------
    pendientes_df = filter_pendientes(filtered_df, source_meta.month_year, source_meta.date)

    try:
        write_pendientes_sheet(output_path, pendientes_df, source_meta.date)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)

    # -------------------------------------------------------------
    # 7. WRITE Semana table: (re)build the Semana/PO Pendiente/
    #    Compromiso/Oportunidad table on 'Resumen - Contrato',
    #    positioned past the pivot's current Total column, showing
    #    every Semana that belongs to the source's month
    # -------------------------------------------------------------
    month_weeks = semanas_in_month(source_meta.month_year)

    try:
        write_semana_table(output_path, month_weeks, source_semana_number, source_meta.month_year)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)

    # -------------------------------------------------------------
    # 8. WRITE chart window: shift the chart-feeding block (rows
    #    18-23) on 'Resumen - Contrato' to the trailing months
    #    ending at the source's current month, without touching
    #    the chart itself (it reads a fixed cell range)
    # -------------------------------------------------------------
    source_year = int(source_meta.month_year.split("-")[1])
    last_data_month = latest_delivery_month(filtered_df, source_year)

    try:
        write_chart_window(output_path, source_meta.month_year, last_data_month)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)

    # -------------------------------------------------------------
    # 9. BUILD PIVOTS: inject fresh pivot tables (Resumen sheets)
    #    cloned from the known-good reference file, based on this
    #    output's own 'Data' sheet, set to auto-refresh on open
    # -------------------------------------------------------------
    rebuild_pivots_from_reference(output_path)

    # -------------------------------------------------------------
    # 10. REPORT
    # -------------------------------------------------------------
    print(f"Semana: {source_meta.semana}")
    print(f"Date: {source_meta.date}")
    print(f"Month-Year: {source_meta.month_year}")
    print(f"Source rows: {len(df)}")
    print(f"Filtered rows for '{args.vendor_name}': {len(filtered_df)}")
    print(filtered_df.head())
    print(f"Pendientes - Contrato rows: {len(pendientes_df)}")
    print(f"Semanas in {source_meta.month_year}: {month_weeks}")
    print(f"Output template: {output_template.name}")
    print(f"Saved output to: {output_path.name}")
