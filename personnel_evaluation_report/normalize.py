"""Explicit, auditable text normalization rules. No fuzzy matching anywhere."""
from __future__ import annotations

import unicodedata

_FA_DIGITS = "۰۱۲۳۴۵۶۷۸۹"
_AR_DIGITS = "٠١٢٣٤٥٦٧٨٩"
_TO_ASCII = str.maketrans(_FA_DIGITS + _AR_DIGITS, "0123456789" * 2)
_TO_FA = str.maketrans("0123456789", _FA_DIGITS)

# Arabic letter variants commonly typed on Arabic keyboards → Persian letters.
_LETTERS = str.maketrans({"ي": "ی", "ى": "ی", "ك": "ک"})

# Invisible characters that sneak in from copy/paste (bidi marks, BOM, NBSP handled separately).
_INVISIBLE = dict.fromkeys(map(ord, "‎‏‪‫‬‭‮⁦⁧⁨⁩﻿"))

# Dash-like characters accepted as equivalent in *headers only*.
_DASHES = str.maketrans({"ـ": "-", "–": "-", "—": "-", "‐": "-", "−": "-"})


def to_ascii_digits(text: str) -> str:
    return text.translate(_TO_ASCII)


def to_fa_digits(text: str) -> str:
    return text.translate(_TO_FA)


def clean_text(value: object) -> str:
    """Cell/display text: NFC, Persian letters, strip invisible marks, collapse whitespace."""
    if value is None:
        return ""
    s = unicodedata.normalize("NFC", str(value))
    s = s.translate(_INVISIBLE).translate(_LETTERS).replace(" ", " ")
    return " ".join(s.split())


def header_key(value: object) -> str:
    """Comparison key for column headers.

    Rules (all explicit): clean_text + every dash variant → '-' with no spaces around it.
    Example: «محل خدمت ـ اداره کل» and «محل خدمت - اداره کل» → same key.
    """
    s = clean_text(value).translate(_DASHES)
    return "-".join(part.strip() for part in s.split("-"))


def normalize_national_id(value: object) -> str:
    """Search/lookup key: digits only (ASCII), leading zeros kept. May return ''."""
    s = clean_text(value)
    s = to_ascii_digits(s)
    for ch in " -_.‌":
        s = s.replace(ch, "")
    return s
