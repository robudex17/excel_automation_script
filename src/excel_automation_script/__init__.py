import argparse
import sys

from .pipeline import PipelineError, run_pipeline


def main() -> None:
    parser = argparse.ArgumentParser(description="Excel automation script")
    parser.add_argument("vendor_name", help="Vendor name to filter the source data by")
    args = parser.parse_args()

    try:
        result = run_pipeline(args.vendor_name, progress=print)
    except PipelineError as e:
        print(f"Error: {e}")
        sys.exit(1)

    print(f"Semana: {result.semana}")
    print(f"Date: {result.date}")
    print(f"Month-Year: {result.month_year}")
    print(f"Source rows: {result.source_rows}")
    print(f"Filtered rows for '{result.vendor_name}': {result.filtered_rows}")
    print(f"Pendientes - Contrato rows: {result.pendientes_rows}")
    print(f"Semanas in {result.month_year}: {result.semanas_in_month}")
    print(f"Output template: {result.output_template_name}")
    print(f"Saved output to: {result.output_path.name}")
