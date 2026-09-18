"""Pure data/analysis engine for the claim dashboard UI."""
from collections import Counter
import re


def clean(value):
    return "" if value is None else str(value).strip()


def month_key(value):
    text = clean(value)
    match = re.search(r"(20\d{2})[/-](\d{1,2})", text)
    if match:
        return f"{match.group(1)[2:]}.{int(match.group(2))}"
    match = re.search(r"(20\d{2})(\d{2})", text)
    return f"{match.group(1)[2:]}.{int(match.group(2))}" if match else ""


class ClaimDataEngine:
    """Keeps DATA operations independent from Tkinter widgets."""

    def __init__(self, headers=None, rows=None):
        self.headers = list(headers or [])
        self.all_rows = [list(row) for row in (rows or [])]

    def set_data(self, headers, rows):
        self.headers = list(headers or [])
        self.all_rows = [list(row) for row in (rows or [])]

    def filter_rows(self, company="전체", model="전체", market="전체", parts=None, names=None):
        parts, names = set(parts or ()), set(names or ())
        return [row for row in self.all_rows
                if (company == "전체" or (len(row) > 0 and clean(row[0]) == company))
                and (model == "전체" or (len(row) > 45 and clean(row[45]) == model))
                and (market == "전체" or (len(row) > 1 and clean(row[1]) == market))
                and (not parts or (len(row) > 11 and clean(row[11]) in parts))
                and (not names or (len(row) > 41 and clean(row[41]) in names))]

    def options(self, index, rows=None):
        source = self.all_rows if rows is None else rows
        return sorted({clean(row[index]) for row in source if len(row) > index and clean(row[index])})

    def counter(self, index, rows):
        return Counter(clean(row[index]) for row in rows if len(row) > index and clean(row[index]))

    def occurrence_by_month(self, rows):
        return Counter(month_key(clean(row[2])[:6]) for row in rows if len(row) > 2 and clean(row[2])[:6])

    def production_by_month(self, rows):
        return Counter(month_key(row[32]) for row in rows if len(row) > 32 and month_key(row[32]))
