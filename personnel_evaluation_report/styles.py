"""Per-report table appearance: text colour, background colour and text style.

Each table section has two parts — «سربرگ» (title bar, column headings, row labels) and
«بدنه» (data cells) — styled independently. Only differences from the default look are stored.
The same values drive the edit-form preview, the report view, the browser print and Excel.
"""
from __future__ import annotations

import re
from dataclasses import asdict, dataclass

# Stable keys for the styleable sections, by official title (order = report order).
SECTION_KEYS: dict[str, str] = {
    "اطلاعات فردی": "personal",
    "اطلاعات پرسنلی": "employment",
    "ارزیابی توانمندی وشایستگی": "competency",
    "کارنامه نظم وانضباط اداری": "discipline",
    "جمع بندی وپیشنهادات مدیریتی": "recommendations",
    "سوابق ارزیابی ۵ سال گذشته": "history",
    "رضایت سنجی": "satisfaction",
    "جمع بندی نهائی ارزیابی عملکرد": "final_summary",
    "نمره نهایی و امضا": "signature",
}
PARTS = ("header", "body")
PART_LABELS = {"header": "سربرگ", "body": "بدنه"}
SIZES = {"small": ("کوچک", 0.85), "normal": ("معمولی", 1.0), "large": ("بزرگ", 1.2)}
_HEX = re.compile(r"^#[0-9a-fA-F]{6}$")


@dataclass(frozen=True)
class PartStyle:
    fg: str
    bg: str
    bold: bool
    italic: bool = False
    underline: bool = False
    size: str = "normal"

    @property
    def scale(self) -> float:
        return SIZES[self.size][1]


DEFAULTS = {
    "header": PartStyle(fg="#1d1d28", bg="#c9c9dd", bold=True),
    "body": PartStyle(fg="#1d1d28", bg="#ffffff", bold=False),
}

# report.styles: {section_key: {part: PartStyle}} — only non-default parts are kept.
Styles = dict[str, dict[str, PartStyle]]


def get(styles: Styles, key: str, part: str) -> PartStyle:
    return styles.get(key, {}).get(part, DEFAULTS[part])


def is_custom(styles: Styles, key: str, part: str) -> bool:
    return part in styles.get(key, {})


def field_name(key: str, part: str, prop: str) -> str:
    return f"st__{key}__{part}__{prop}"


def parse_form(form) -> tuple[Styles, list[str]]:
    """Read every section's style controls. Invalid values are reported, not guessed."""
    styles: Styles = {}
    errors: list[str] = []
    for title, key in SECTION_KEYS.items():
        for part in PARTS:
            d = DEFAULTS[part]
            fg = (form.get(field_name(key, part, "fg")) or d.fg).strip()
            bg = (form.get(field_name(key, part, "bg")) or d.bg).strip()
            size = form.get(field_name(key, part, "size")) or d.size
            if not (_HEX.match(fg) and _HEX.match(bg)) or size not in SIZES:
                errors.append(f"تنظیمات ظاهر «{title}» نامعتبر بود و نادیده گرفته شد.")
                continue
            # Checkboxes: a hidden "present" marker tells "unchecked" apart from "control not on the page".
            present = form.get(field_name(key, part, "present"))
            def flag(prop: str, default: bool) -> bool:
                return bool(form.get(field_name(key, part, prop))) if present else default
            ps = PartStyle(fg=fg.lower(), bg=bg.lower(), bold=flag("bold", d.bold),
                           italic=flag("italic", d.italic), underline=flag("underline", d.underline), size=size)
            if ps != d:
                styles.setdefault(key, {})[part] = ps
    return styles, errors


def css(styles: Styles) -> str:
    """Stylesheet for the report view and edit form (served from our own origin — CSP-safe)."""
    out = []
    for key, parts in styles.items():
        sel = f'section.rsec[data-sec="{key}"]'   # beats the generic table rules in app.css
        for part, ps in parts.items():
            if part == "header":
                targets = (f"{sel} h3", f"{sel} thead th", f"{sel} table.kv th")
            else:
                targets = (f"{sel} tbody td", f"{sel} table.kv td", f"{sel} .note")
            decl = (f"color:{ps.fg};background:{ps.bg};font-weight:{700 if ps.bold else 400};"
                    f"font-style:{'italic' if ps.italic else 'normal'};"
                    f"text-decoration:{'underline' if ps.underline else 'none'}")
            out.append(f"{', '.join(targets)} {{{decl}}}")
            if ps.scale != 1:
                cells = [t for t in targets if not t.endswith(" h3")]
                out.append(f"{', '.join(cells)} {{font-size:{ps.scale}em}}")
                if part == "header":
                    out.append(f"{sel} h3 {{font-size:calc(1.05rem * {ps.scale})}}")
    return "\n".join(out) + "\n"


def to_jsonable(styles: Styles) -> dict:
    return {k: {p: asdict(v) for p, v in parts.items()} for k, parts in styles.items()}
