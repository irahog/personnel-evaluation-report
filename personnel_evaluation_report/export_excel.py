"""Excel output: fill the official template for one employee.

Two modes, both using the same filled sheet:
  * build_single_workbook()  – a new workbook with just this employee's sheet (download)
  * append_to_shared()       – add the sheet to the shared «کارنامه» workbook on disk

The source personnel workbook is never touched here.
"""
from __future__ import annotations

import io
import json
import re
import os
import tempfile
import zipfile
from copy import copy
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries
from openpyxl.styles import Font
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.worksheet import Worksheet

from . import fonts, photo as photo_mod
from .normalize import clean_text, normalize_national_id
from .paths import TEMPLATE_MAP, TEMPLATE_XLSX
from .report import FIELDS, MARITAL_MARRIED, MARITAL_SINGLE, Report

_BAD_SHEET_CHARS = str.maketrans({c: " " for c in '[]:*?/\\'})


class ExportError(Exception):
    """Persian, user-facing; contains no personal data."""


@dataclass
class FontDecision:
    used: str
    warning: str = ""


def decide_font(chosen: str, fallback: str) -> FontDecision:
    """Use the chosen font if Excel can render it here, otherwise the configured fallback (with a warning)."""
    if fonts.is_installed(chosen):
        return FontDecision(chosen)
    if fallback and fallback != chosen and fonts.is_installed(fallback):
        return FontDecision(fallback, f"فونت «{chosen}» روی این رایانه نصب نیست؛ در فایل اکسل از فونت جایگزین «{fallback}» استفاده شد.")
    return FontDecision(chosen, f"فونت «{chosen}» روی این رایانه نصب نیست؛ نام آن در فایل اکسل ثبت شد ولی اکسل تا نصب فونت، "
                                "متن را با فونت دیگری نمایش می‌دهد.")


def _template_map() -> dict[str, str]:
    return json.loads(TEMPLATE_MAP.read_text(encoding="utf-8"))["fields"]


def _image_boxes() -> dict[str, str]:
    return json.loads(TEMPLATE_MAP.read_text(encoding="utf-8")).get("images", {})


EMU_PER_PX = 9525


def _col_px(ws: Worksheet, col: int) -> float:
    w = ws.column_dimensions[get_column_letter(col)].width or 8.43
    return w * 7 + 5                      # Excel's character-width → pixel rule (default font)


def _row_px(ws: Worksheet, row: int) -> float:
    h = ws.row_dimensions[row].height or 15
    return h * 96 / 72                    # points → pixels


def add_photo(ws: Worksheet, report: Report) -> None:
    """Place the employee photo inside its merged box, fitted and centred — never stretched."""
    box = _image_boxes().get("photo")
    if not box or not report.photo:
        return
    min_col, min_row, max_col, max_row = range_boundaries(box)
    box_w = sum(_col_px(ws, c) for c in range(min_col, max_col + 1))
    box_h = sum(_row_px(ws, r) for r in range(min_row, max_row + 1))
    x, y, w, h = photo_mod.fit(box_w, box_h, report.photo.width, report.photo.height, pad=6)
    # Convert the in-box offset to (cell, offset-in-cell) as Excel's anchor requires.
    col, off_x = min_col, x
    while col < max_col and off_x >= _col_px(ws, col):
        off_x -= _col_px(ws, col)
        col += 1
    row, off_y = min_row, y
    while row < max_row and off_y >= _row_px(ws, row):
        off_y -= _row_px(ws, row)
        row += 1
    img = XLImage(io.BytesIO(report.photo.jpeg))
    img.anchor = OneCellAnchor(
        _from=AnchorMarker(col=col - 1, colOff=int(off_x * EMU_PER_PX), row=row - 1, rowOff=int(off_y * EMU_PER_PX)),
        ext=XDRPositiveSize2D(int(w * EMU_PER_PX), int(h * EMU_PER_PX)))
    ws.add_image(img)


def sheet_title(name: str) -> str:
    t = clean_text(name).translate(_BAD_SHEET_CHARS).strip(" '")
    return " ".join(t.split())[:31] or "کارمند"


def fill_sheet(ws: Worksheet, report: Report, font_name: str) -> None:
    cells = _template_map()
    for fid, f in FIELDS.items():
        if f.input == "marital":
            ws[cells["marital_married_mark"]].value = "☑" if report.marital == MARITAL_MARRIED else "☐"
            ws[cells["marital_single_mark"]].value = "☑" if report.marital == MARITAL_SINGLE else "☐"
            continue
        addr = cells[fid]
        if f.input == "check":
            ws[addr].value = "☑" if report.manual.get(fid) else "☐"
        elif f.input in ("number", "count"):
            ws[addr].value = report.number(fid)
        else:
            ws[addr].value = report.text(fid) or None
    if report.photo:
        ws[cells["photo"]].value = None     # placeholder text only shows when there is no photo
    # One font for the whole report, except symbol cells (☐ ★ …) that need a symbol font.
    for row in ws.iter_rows():
        for c in row:
            if c.font and c.font.name != fonts.SYMBOL_FONT:
                f = copy(c.font)
                c.font = Font(name=font_name, size=f.size, bold=f.bold, italic=f.italic,
                              color=f.color, underline=f.underline)


def _copy_sheet(src: Worksheet, dst: Worksheet) -> None:
    """Copy values, styles, merges, sizes, view and print setup between workbooks."""
    for row in src.iter_rows():
        for c in row:
            d = dst.cell(row=c.row, column=c.column, value=c.value)
            if c.has_style:
                d.font, d.fill, d.border = copy(c.font), copy(c.fill), copy(c.border)
                d.alignment, d.number_format, d.protection = copy(c.alignment), c.number_format, copy(c.protection)
    for rng in src.merged_cells.ranges:
        dst.merge_cells(str(rng))
    for k, dim in src.column_dimensions.items():
        dst.column_dimensions[k].width = dim.width
        dst.column_dimensions[k].hidden = dim.hidden
    for k, dim in src.row_dimensions.items():
        dst.row_dimensions[k].height = dim.height
    dst.sheet_view.rightToLeft = src.sheet_view.rightToLeft
    dst.sheet_view.showGridLines = src.sheet_view.showGridLines
    ps, dps = src.page_setup, dst.page_setup
    dps.paperSize, dps.orientation = ps.paperSize, ps.orientation
    dps.fitToWidth, dps.fitToHeight = ps.fitToWidth, ps.fitToHeight
    dst.sheet_properties.pageSetUpPr.fitToPage = src.sheet_properties.pageSetUpPr.fitToPage
    dst.print_options.horizontalCentered = src.print_options.horizontalCentered
    for attr in ("left", "right", "top", "bottom", "header", "footer"):
        setattr(dst.page_margins, attr, getattr(src.page_margins, attr))
    if src.print_area:
        dst.print_area = src.print_area
    for brk in src.row_breaks.brk:
        dst.row_breaks.append(copy(brk))


def _filled_template(report: Report, font_name: str) -> Worksheet:
    tpl = load_workbook(TEMPLATE_XLSX)
    ws = tpl.worksheets[0]
    fill_sheet(ws, report, font_name)
    return ws


def build_single_workbook(report: Report, font_name: str) -> bytes:
    src = _filled_template(report, font_name)
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title(report.employee_name)
    _copy_sheet(src, ws)
    add_photo(ws, report)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Shared workbook ──────────────────────────────────────────────────────

@dataclass
class AppendResult:
    sheet: str
    replaced: bool


class SameEmployeeExists(ExportError):
    def __init__(self, sheet: str):
        super().__init__(f"برای این کارمند قبلاً برگه‌ای با نام «{sheet}» در کارنامهٔ مشترک وجود دارد.")
        self.sheet = sheet


_NON_PICTURE = re.compile(rb"<(?:\w+:)?(?:sp|cxnSp|grpSp|graphicFrame)[\s>]")


def _has_drawings(path: Path) -> bool:
    """True if the workbook has drawings that openpyxl would lose on save.

    Pictures (such as employee photos added by this app) survive a re-save; shapes, charts,
    grouped drawings and form controls do not.
    """
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if any(n.startswith(("xl/charts/", "xl/ctrlProps/")) or n.endswith(".vml") for n in names):
            return True
        return any(_NON_PICTURE.search(z.read(n)) for n in names
                   if n.startswith("xl/drawings/") and n.endswith(".xml"))


def _existing_sheet_for(wb, report: Report) -> str | None:
    addr = _template_map()["national_id"]
    key = normalize_national_id(report.auto.get("national_id", ""))
    for ws in wb.worksheets:
        v = ws[addr].value
        if v is not None and normalize_national_id(v) == key:
            return ws.title
    return None


def _unique_title(wb, base: str) -> str:
    names = {n.lower() for n in wb.sheetnames}
    if base.lower() not in names:
        return base
    for i in range(2, 1000):
        suffix = f" ({i})"
        cand = base[: 31 - len(suffix)] + suffix
        if cand.lower() not in names:
            return cand
    raise ExportError("نام برگهٔ تازه قابل ساخت نیست.")


def append_to_shared(path: str, report: Report, font_name: str, *, replace: bool = False) -> AppendResult:
    p = Path(path)
    if p.suffix.lower() != ".xlsx":
        raise ExportError("کارنامهٔ مشترک باید فایل ‎.xlsx‎ باشد.")
    new_file = not p.exists()
    if new_file:
        wb = Workbook()
        wb.remove(wb.active)
    else:
        try:
            if _has_drawings(p):
                raise ExportError("این کارنامهٔ مشترک شامل شکل، تصویر یا کنترل فرم است و ذخیرهٔ دوباره با این برنامه آن‌ها را از بین می‌برد. "
                                  "لطفاً یک کارنامهٔ مشترک تازه بسازید.")
            wb = load_workbook(p)
        except ExportError:
            raise
        except PermissionError:
            raise ExportError("دسترسی به کارنامهٔ مشترک ممکن نیست. اگر در اکسل باز است، آن را ببندید.")
        except (zipfile.BadZipFile, InvalidFileException, KeyError, ValueError, OSError):
            raise ExportError("کارنامهٔ مشترک خراب است یا قابل خواندن نیست.")

    existing = None if new_file else _existing_sheet_for(wb, report)
    replaced = False
    if existing:
        if not replace:
            raise SameEmployeeExists(existing)
        idx = wb.sheetnames.index(existing)
        wb.remove(wb[existing])
        title = existing
        ws = wb.create_sheet(title, idx)
        replaced = True
    else:
        title = _unique_title(wb, sheet_title(report.employee_name))
        ws = wb.create_sheet(title)
    _copy_sheet(_filled_template(report, font_name), ws)
    add_photo(ws, report)
    wb.active = wb.sheetnames.index(title)
    for other in wb.worksheets:
        other.sheet_view.tabSelected = other.title == title

    _atomic_save(wb, p)
    return AppendResult(title, replaced)


def _atomic_save(wb, p: Path) -> None:
    """Write next to the target, then replace — the shared file is never left half-written."""
    try:
        fd, tmp = tempfile.mkstemp(prefix="~per-report-", suffix=".xlsx", dir=p.parent)
    except OSError:
        raise ExportError("در پوشهٔ کارنامهٔ مشترک امکان نوشتن وجود ندارد.")
    os.close(fd)
    try:
        wb.save(tmp)
        os.replace(tmp, p)
    except PermissionError:
        raise ExportError("ذخیرهٔ کارنامهٔ مشترک ممکن نشد. اگر فایل در اکسل باز است، آن را ببندید و دوباره تلاش کنید.")
    except OSError:
        raise ExportError("ذخیرهٔ کارنامهٔ مشترک با خطا مواجه شد.")
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
