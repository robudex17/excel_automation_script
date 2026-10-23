from calendar import monthrange
from datetime import date, datetime

# -------------------------------------------------------------
# CONFIGURATION
# -------------------------------------------------------------
MONTH_YEAR_FORMAT = "%b-%Y"


# -------------------------------------------------------------
# COMPUTE: every ISO week number any day of the given month
# touches, in chronological order. A month spans 4 or 5 weeks
# depending on how its start/end align with week boundaries.
# Matches the source filename's own 'Semana N' convention
# (ISO week, Monday-start) -- NOT the Data sheet's per-row
# SEMANA column, which uses Excel's WEEKNUM(date) default
# (Sunday-start) on each PO's own delivery date.
# -------------------------------------------------------------
def semanas_in_month(month_year: str) -> list[int]:
    parsed = datetime.strptime(month_year, MONTH_YEAR_FORMAT)
    year, month = parsed.year, parsed.month
    _, days_in_month = monthrange(year, month)

    weeks: list[int] = []
    for day in range(1, days_in_month + 1):
        week = date(year, month, day).isocalendar()[1]
        if week not in weeks:
            weeks.append(week)
    return weeks


# -------------------------------------------------------------
# VALIDATE: the source's own Semana must actually fall within
# the month its own filename claims, e.g. 'Semana 38' must be
# one of the weeks that 'Sep-2026' spans
# -------------------------------------------------------------
def validate_semana_in_month(semana: int, month_year: str) -> None:
    valid_weeks = semanas_in_month(month_year)
    if semana not in valid_weeks:
        raise ValueError(
            f"Semana {semana} does not belong to {month_year} (valid weeks: {valid_weeks})"
        )
