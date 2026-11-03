from calendar import monthrange
from datetime import date, datetime, timedelta

# -------------------------------------------------------------
# CONFIGURATION
# -------------------------------------------------------------
MONTH_YEAR_FORMAT = "%b-%Y"
REPORT_DATE_FORMAT = "%d%b%Y"
ROLLING_WINDOW_SIZE = 5


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
# COMPUTE: a forward-looking rolling window of ROLLING_WINDOW_SIZE
# (5) consecutive ISO weeks, starting at the report's own Semana --
# a "running result" projection (per the client's own process: each
# report is generated once and sent, not reopened later), not
# bounded by the calendar month. A report near month-end naturally
# stretches into the following month's weeks, e.g. Semana 44 (the
# last week of October) continues into 45-48 (November).
#
# Computed from the report's own real date (not simple integer
# week-number arithmetic), so the ISO year-end wraparound (week 52
# or 53 rolling over to week 1 of the next year) is handled
# correctly by the calendar itself rather than risking an
# off-by-one near a year boundary.
# -------------------------------------------------------------
def rolling_semanas(report_date: str, window_size: int = ROLLING_WINDOW_SIZE) -> list[int]:
    anchor = datetime.strptime(report_date, REPORT_DATE_FORMAT).date()
    monday = anchor - timedelta(days=anchor.isoweekday() - 1)
    return [(monday + timedelta(weeks=i)).isocalendar()[1] for i in range(window_size)]


# -------------------------------------------------------------
# DECIDE: which week list the Semana table actually shows --
# the full calendar month (semanas_in_month) when the forward
# rolling window stays entirely within it, or the forward window
# itself (dropping already-elapsed weeks) when it would stretch
# into the next month. Flips only at the boundary; the "full
# month" case can include weeks before the report's own Semana
# (e.g. Semana 32 of Aug-2026 still shows 31-36, including 31,
# which precedes it) -- that's intentional: it only switches to a
# forward-only, history-dropping view once a month crossing
# actually happens.
# -------------------------------------------------------------
def semana_table_weeks(report_date: str, month_year: str) -> list[int]:
    rolling = rolling_semanas(report_date)
    month_weeks = semanas_in_month(month_year)
    if all(week in month_weeks for week in rolling):
        return month_weeks
    return rolling


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
