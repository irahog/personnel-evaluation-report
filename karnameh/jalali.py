"""Jalali (Persian) date parsing, validation and Gregorian→Jalali conversion.

Parsing, validation and formatting are kept separate. Source dates are Jalali text
«yyyy/mm/dd» (often with Persian digits) and are displayed as-is; this module only
checks they are plausible and converts genuine Excel date cells (Gregorian).
"""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass

from .normalize import to_ascii_digits, to_fa_digits

_DATE_RE = re.compile(r"^(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})$")

# 33-year arithmetic cycle; matches the official calendar for 1178–1634 AP.
_LEAP_REMAINDERS = {1, 5, 9, 13, 17, 22, 26, 30}


@dataclass(frozen=True)
class JalaliDate:
    year: int
    month: int
    day: int

    def format(self) -> str:
        return to_fa_digits(f"{self.year:04d}/{self.month:02d}/{self.day:02d}")


def is_leap(year: int) -> bool:
    return year % 33 in _LEAP_REMAINDERS


def month_length(year: int, month: int) -> int:
    if 1 <= month <= 6:
        return 31
    if 7 <= month <= 11:
        return 30
    if month == 12:
        return 30 if is_leap(year) else 29
    raise ValueError(month)


def parse(text: str) -> JalaliDate | None:
    """Parse «yyyy/mm/dd» (Persian/Arabic/ASCII digits). Returns None when invalid."""
    m = _DATE_RE.match(to_ascii_digits(text.strip()))
    if not m:
        return None
    y, mo, d = (int(g) for g in m.groups())
    if not (1200 <= y <= 1600 and 1 <= mo <= 12 and 1 <= d <= month_length(y, mo)):
        return None
    return JalaliDate(y, mo, d)


def from_gregorian(date: dt.date) -> JalaliDate:
    """Gregorian → Jalali by day counting from a known epoch (1 Farvardin 1300 = 21 Mar 1921)."""
    days = (date - dt.date(1921, 3, 21)).days
    if days < 0:
        raise ValueError("date before 1300 AP is not supported")
    year = 1300
    while True:
        ylen = 366 if is_leap(year) else 365
        if days < ylen:
            break
        days -= ylen
        year += 1
    month = 1
    while days >= month_length(year, month):
        days -= month_length(year, month)
        month += 1
    return JalaliDate(year, month, days + 1)
