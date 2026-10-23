import re
from datetime import date as date_cls
from pathlib import Path
from typing import NamedTuple

# -------------------------------------------------------------
# CONFIGURATION
# -------------------------------------------------------------
# Matches source filenames like:
#   "...Semana 27 - 30Jun2026"        -> semana="27", day="30", month="Jun",        year="2026"
#   "...Semana 38 - 14September2026"  -> semana="38", day="14", month="September",  year="2026"
#   "...Semana 38 - 14Septiembre2026" -> semana="38", day="14", month="Septiembre", year="2026"
FILENAME_PATTERN = re.compile(
    r"Semana\s+(?P<semana>\d+)\s*-\s*(?P<day>\d{1,2})(?P<month>[A-Za-z]+)(?P<year>\d{4})"
)
OUTPUT_DATE_FORMAT = "%d%b%Y"
MONTH_YEAR_FORMAT = "%b-%Y"

# Accepts English or Spanish, abbreviated or full month names, case-insensitive.
MONTH_NAME_TO_NUMBER = {
    "jan": 1, "january": 1, "ene": 1, "enero": 1,
    "feb": 2, "february": 2, "febrero": 2,
    "mar": 3, "march": 3, "marzo": 3,
    "apr": 4, "april": 4, "abr": 4, "abril": 4,
    "may": 5, "mayo": 5,
    "jun": 6, "june": 6, "junio": 6,
    "jul": 7, "july": 7, "julio": 7,
    "aug": 8, "august": 8, "ago": 8, "agosto": 8,
    "sep": 9, "sept": 9, "september": 9, "set": 9, "septiembre": 9, "setiembre": 9,
    "oct": 10, "october": 10, "octubre": 10,
    "nov": 11, "november": 11, "noviembre": 11,
    "dec": 12, "december": 12, "dic": 12, "diciembre": 12,
}


class SourceMetadata(NamedTuple):
    semana: str      # e.g. "Semana 27"
    date: str         # e.g. "30Jun2026" -- always normalized to DDMonYYYY (English abbreviation)
    month_year: str    # e.g. "Jun-2026" (date with the day omitted)


# -------------------------------------------------------------
# PARSE: pull Semana/date/month-year out of the source filename,
# accepting English or Spanish, abbreviated or full month names
# -------------------------------------------------------------
def parse_source_filename(path: Path) -> SourceMetadata:
    match = FILENAME_PATTERN.search(path.stem)
    if not match:
        raise ValueError(
            f"Filename does not match expected pattern 'Semana <n> - DDMonthYYYY': {path.name}"
        )

    semana = f"Semana {match.group('semana')}"

    month_number = MONTH_NAME_TO_NUMBER.get(match.group("month").lower())
    if month_number is None:
        raise ValueError(
            f"Unrecognized month '{match.group('month')}' in filename: {path.name}"
        )

    parsed_date = date_cls(int(match.group("year")), month_number, int(match.group("day")))

    date_str = parsed_date.strftime(OUTPUT_DATE_FORMAT)
    month_year = parsed_date.strftime(MONTH_YEAR_FORMAT)

    return SourceMetadata(semana=semana, date=date_str, month_year=month_year)
