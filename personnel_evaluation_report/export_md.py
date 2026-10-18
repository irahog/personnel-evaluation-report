"""Markdown output, generated from the same Report + SECTIONS as every other output."""
from __future__ import annotations

import base64

from .normalize import to_fa_digits
from .report import FIELDS, MARITAL_LABELS, Banner, FieldsSection, Report, SECTIONS, Static, TableSection


def _md(text: str) -> str:
    """Escape for a Markdown table cell / line."""
    text = text.replace("\\", "\\\\").replace("|", "\\|").replace("*", "\\*").replace("_", "\\_")
    return text.replace("\n", "<br>")   # only HTML used: line breaks inside table cells


def field_value(report: Report, fid: str) -> str:
    if FIELDS[fid].input == "marital":
        return "  ".join(("☑ " if report.marital == k else "☐ ") + v for k, v in MARITAL_LABELS.items())
    return report.text(fid)


def render(report: Report) -> str:
    lines = ["# کارنامه پرسنلی", ""]
    for sec in SECTIONS:
        if isinstance(sec, Banner):
            lines += [f"## {sec.title}", ""]
        elif isinstance(sec, FieldsSection):
            lines += [f"### {sec.title}", ""]
            if sec.photo and report.photo:
                b64 = base64.b64encode(report.photo.jpeg).decode("ascii")
                lines += [f"![عکس کارمند](data:image/jpeg;base64,{b64})", ""]
            lines += ["| عنوان | مقدار |", "|---|---|"]
            for fid in sec.fields:
                lines.append(f"| {_md(FIELDS[fid].label)} | {_md(field_value(report, fid))} |")
            lines.append("")
        elif isinstance(sec, TableSection):
            lines += [f"### {sec.title}", "",
                      "| " + " | ".join(map(_md, sec.columns)) + " |",
                      "|" + "---|" * len(sec.columns)]
            for row in sec.rows:
                cells = [c.text if isinstance(c, Static) else field_value(report, c) for c in row]
                cells += [""] * (len(sec.columns) - len(cells))
                lines.append("| " + " | ".join(_md(c) for c in cells) + " |")
            if sec.note:
                lines += ["", _md(sec.note)]
            lines.append("")
    lines += [
        "---",
        "",
        f"منبع اطلاعات: فایل «{_md(report.source_file)}»، برگهٔ «{_md(report.source_sheet)}»، "
        f"ردیف {to_fa_digits(str(report.source_row))} — زمان تهیه: {report.created_fa}",
        "",
    ]
    return "\n".join(lines)
