import pandas as pd

# -------------------------------------------------------------
# CONFIGURATION
# -------------------------------------------------------------
VENDOR_COLUMN = "Vendor(VN) - Text"

CALC_DELIVERY_COLUMN = "Date : Calculated Delivery - Key"
GR_POSTING_COLUMN = "Date: First GR Posting - Key"
DELIVERY_COMPLETION_COLUMN = "Indicator : Delivery Completion - Key"
CONTRACT_COLUMN = "Contract (Outline Agreement) No - Key"


# -------------------------------------------------------------
# FILTER: keep only rows for the requested vendor (as a copy, so
# the original dataframe from io.load_source() stays untouched)
# -------------------------------------------------------------
def filter_by_vendor(df: pd.DataFrame, vendor_name: str) -> pd.DataFrame:
    return df[df[VENDOR_COLUMN] == vendor_name].copy()


# -------------------------------------------------------------
# HELPER: mirrors the 'Data' sheet's DELIVERY formula in pandas
# -------------------------------------------------------------
def _delivery_status(row: pd.Series) -> str:
    gr_post = str(row[GR_POSTING_COLUMN]).strip()
    del_comp = str(row[DELIVERY_COMPLETION_COLUMN]).strip()
    if gr_post != "#" and del_comp == "X":
        return "COMPLETA"
    if gr_post != "#" and del_comp != "X":
        return "PARCIAL"
    return "PENDIENTE"


# -------------------------------------------------------------
# FILTER: keep only rows destined for 'Pendientes - Contrato':
# an active contract, due in the source's month, not yet complete
# (PENDIENTE or PARCIAL), and not yet overdue (PERFORMANCE = OK).
# Sorted by delivery date. "Overdue" is judged against the source
# report's own date (as_of_date), not the real-world clock --
# otherwise reprocessing an older report always shows everything
# as overdue.
# -------------------------------------------------------------
PENDIENTES_DELIVERY_STATUSES = ["PENDIENTE", "PARCIAL"]


def filter_pendientes(df: pd.DataFrame, month_year: str, as_of_date: str) -> pd.DataFrame:
    calc_date = pd.to_datetime(df[CALC_DELIVERY_COLUMN], errors="coerce")
    delivery = df.apply(_delivery_status, axis=1)
    mes = calc_date.dt.strftime("%b-%Y")
    today = pd.to_datetime(as_of_date, format="%d%b%Y").normalize()
    performance = (calc_date >= today).map({True: "OK", False: "NOK"})

    mask = (
        (df[CONTRACT_COLUMN].astype(str).str.strip() != "#")
        & delivery.isin(PENDIENTES_DELIVERY_STATUSES)
        & (mes == month_year)
        & (performance == "OK")
    )

    result = df[mask].copy()
    result = result.assign(_sort_date=calc_date[mask])
    result = result.sort_values("_sort_date").drop(columns="_sort_date")
    return result


# -------------------------------------------------------------
# FIND: the latest calendar month (1-12) within the given year
# that this vendor actually has any delivery date scheduled for.
# Vendors legitimately differ here -- e.g. one vendor's real PO
# commitments might only extend through October while another's
# reach December -- so this isn't a fixed/shared value.
# -------------------------------------------------------------
def latest_delivery_month(df: pd.DataFrame, year: int) -> int:
    calc_date = pd.to_datetime(df[CALC_DELIVERY_COLUMN], errors="coerce")
    this_year = calc_date[calc_date.dt.year == year]
    if this_year.empty:
        return 0
    return int(this_year.dt.month.max())
