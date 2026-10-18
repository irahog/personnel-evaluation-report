"""Build the blank «کارنامه پرسنلی» report template from the official layout.

Layout is reconstructed from photos of the colleague's workbook (FinalExcel_1..3).
Output:
  {Template,personnel_evaluation_report/resources}/report_template.xlsx      – one blank sheet named «قالب»
  {Template,personnel_evaluation_report/resources}/report_template_map.json  – field id -> cell address

Run:  .venv/bin/python tools/build_template.py
"""
from __future__ import annotations

import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.worksheet import Worksheet

ROOT = Path(__file__).resolve().parent.parent
OUT_DIRS = (ROOT / "Template", ROOT / "personnel_evaluation_report" / "resources")

FONT = "B Nazanin"
SYMBOL_FONT = "Segoe UI Symbol"  # ☐ ☑ ★ … are not in the Persian fonts
SHEET = "قالب"

# Colours approximated from the photos.
C_TITLE = "A9A9C8"
C_HEAD = "C9C9DD"
C_SUB = "E4E4EF"
C_YELLOW = "FFFF00"
C_GIRL = "F4B6C2"
C_BOY = "9DC3E6"
C_WARN = "FFC000"
C_PLEDGE = "C00000"
C_PEACH = "F8CBAD"
C_SCORE = "F4B183"

thin = Side(style="thin", color="000000")
medium = Side(style="medium", color="000000")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

field_map: dict[str, str] = {}
image_boxes: dict[str, str] = {}   # picture id -> merged cell range


def fill(color: str) -> PatternFill:
    return PatternFill("solid", start_color=color, end_color=color)


def cell(ws: Worksheet, rng: str, value=None, *, bold=False, size=12, color="000000",
         bg: str | None = None, h="center", v="center", field: str | None = None,
         font: str = FONT) -> None:
    """Merge `rng`, write value and style every cell in the range (borders need it)."""
    if ":" in rng:
        ws.merge_cells(rng)
    first = rng.split(":")[0]
    ws[first].value = value
    for row in ws[rng] if ":" in rng else [[ws[rng]]]:
        for c in row:
            c.font = Font(name=font, size=size, bold=bold, color=color)
            c.alignment = Alignment(horizontal=h, vertical=v, wrap_text=True, readingOrder=2,
                                    indent=1 if h == "right" else 0)
            c.border = BOX
            if bg:
                c.fill = fill(bg)
    if field:
        assert field not in field_map, field
        field_map[field] = first


def build() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False

    widths = {"I": 13, "J": 11}
    for col in "BCDEFGHIJKLMNOPQRSUVWX":
        ws.column_dimensions[col].width = widths.get(col, 10)
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["T"].width = 2
    ws.column_dimensions["Y"].width = 2
    for r in range(1, 50):
        ws.row_dimensions[r].height = 30 if r <= 21 else 26

    # ── 1. Title ─────────────────────────────────────────────────────────
    cell(ws, "B1:X2", "کارنامه پرسنلی", bold=True, size=18, bg=C_TITLE)

    # ── 2. اطلاعات فردی (B–M, photo box B–D) / 3. اطلاعات پرسنلی (N–X) ─────
    cell(ws, "B4:M5", "اطلاعات فردی", bold=True, size=14, bg=C_HEAD)
    cell(ws, "N4:X5", "اطلاعات پرسنلی", bold=True, size=14, bg=C_HEAD)

    # Employee photo box, right of the personal-info table (RTL: column B is the rightmost).
    # The app places the picture inside it, fitted without distortion; the text is a placeholder.
    cell(ws, "B6:D19", "عکس کارمند", size=11, color="9A9AB0", field="photo")
    image_boxes["photo"] = "B6:D19"

    personal_rows = [6, 8, 10, 12, 14, 16, 18]
    individual = [
        ("full_name", "نام ونام خانوادگی"),
        ("national_id", "کدملی:"),
        ("personnel_no", "کدپرسنلی"),
        ("birth_date", "تاریخ تولد:"),
        ("degree", "مدرک تحصیلی"),
        ("field_of_study", "رشته تحصیلی"),
        ("mobile", "شماره تماس:"),
    ]
    for r, (fid, label) in zip(personal_rows, individual):
        cell(ws, f"E{r}:G{r+1}", label, bold=True, h="right")
        cell(ws, f"H{r}:K{r+1}", None, field=fid)

    # Marital status: two boxes side by side, tick under the matching one.
    cell(ws, "L6:L7", "متأهل", bold=True)
    cell(ws, "M6:M7", "مجرد", bold=True)
    cell(ws, "L8:L9", "☐", size=16, field="marital_married_mark", font=SYMBOL_FONT)
    cell(ws, "M8:M9", "☐", size=16, field="marital_single_mark", font=SYMBOL_FONT)
    # Children: daughter / son, entered manually.
    cell(ws, "L10:M11", "تعدادفرزندان", bold=True)
    cell(ws, "L12:L13", "دختر", bold=True, bg=C_GIRL)
    cell(ws, "M12:M13", "پسر", bold=True, bg=C_BOY)
    cell(ws, "L14:L19", None, field="children_daughters")
    cell(ws, "M14:M19", None, field="children_sons")

    employment = [
        ("employment_type", "نوع استخدام"),
        ("post_title", "پست سازمانی"),
        ("work_location", "محل خدمت"),
        ("start_date", "تاریخ شروع به کار"),
        ("service_history", "سوابق خدمتی/تجربی"),
        ("training_history", "سوابق آموزشی"),
    ]
    for r, (fid, label) in zip(personal_rows, employment):
        cell(ws, f"N{r}:P{r+1}", label, bold=True)
        cell(ws, f"Q{r}:T{r+1}", None, field=fid)
    cell(ws, "N18:P19", "سوابق ارزیابی", bold=True)
    cell(ws, "Q18:R18", "تنبیهات", bold=True, size=11, bg=C_SUB)
    cell(ws, "S18:T18", "تشویق ها", bold=True, size=11, bg=C_SUB)
    cell(ws, "Q19:R19", None, field="eval_punishments")
    cell(ws, "S19:T19", None, field="eval_rewards")

    # Yellow box: previous posts / post suspension.
    cell(ws, "U6:X6", "سمت های قبلی:", bold=True, size=14, bg=C_YELLOW, h="right")
    cell(ws, "U7:X13", None, bg=C_YELLOW, h="right", v="top", field="previous_posts")
    cell(ws, "U14:X14", "توقف پست:", bold=True, size=14, bg=C_YELLOW, h="right")
    cell(ws, "U15:X19", None, bg=C_YELLOW, h="right", v="top", field="post_suspension")

    # ── 4. Banner ────────────────────────────────────────────────────────
    cell(ws, "B20:X21", "عملکردوتوانمندی کارکنان", bold=True, size=14, bg=C_HEAD)

    # ── 5. ارزیابی توانمندی وشایستگی (B–I) ────────────────────────────────
    cell(ws, "B23:I23", "ارزیابی توانمندی وشایستگی", bold=True, size=13, bg=C_HEAD)
    for rng, t in (("B24", "ردیف"), ("C24:E24", "شاخص"), ("F24", "نمره"), ("G24:I24", "شرح")):
        cell(ws, rng, t, bold=True, bg=C_SUB)
    competencies = [
        "تخصص وتسلط بروظایف", "دقت وکیفیت انجام کار", "مسئولیت پذیری",
        "توانائی حل مسئله ومدیریت امور", "همکاری وکارتیمی", "تعامل وارتباط مؤثر",
        "ابتکاروارائه پیشنهادموثر",
    ]
    for i, label in enumerate(competencies, 1):
        r = 24 + i
        cell(ws, f"B{r}", _fa(i))
        cell(ws, f"C{r}:E{r}", label)
        cell(ws, f"F{r}", None, field=f"comp_{i}_score")
        cell(ws, f"G{r}:I{r}", None, size=11, field=f"comp_{i}_desc")
    cell(ws, "B32:E33", "جمع امتیاز", bold=True, bg=C_HEAD)
    cell(ws, "F32:F33", None, bold=True, field="comp_total")
    cell(ws, "G32:I33", None, bg=C_HEAD)

    # ── 6. کارنامه نظم وانضباط اداری (K–S) ───────────────────────────────
    cell(ws, "K23:S23", "کارنامه نظم وانضباط اداری", bold=True, size=13, bg=C_HEAD)
    for rng, t in (("K24", "ردیف"), ("L24:N24", "شاخص"), ("O24", "نمره"),
                   ("P24:S24", "توضیحات/مستندات")):
        cell(ws, rng, t, bold=True, bg=C_SUB)
    discipline = [
        "ورود وخروج به موقع", "میزان تأخیر/غیبت غیرموجه", "رعایت شئونات اداری",
        "پایبندی به اصول اخلاق حرفه ای", "آراستگی وپوشش", "شرکت درمراسم مذهبی_فرهنگی",
    ]
    for i, label in enumerate(discipline, 1):
        r = 24 + i
        cell(ws, f"K{r}", _fa(i))
        cell(ws, f"L{r}:N{r}", label)
        cell(ws, f"O{r}", None, field=f"disc_{i}_score")
        cell(ws, f"P{r}:S{r}", None, size=11, field=f"disc_{i}_desc")
    cell(ws, "K31:K33", _fa(7))
    cell(ws, "L31:N33", "تعهدات وتذکرات")
    cell(ws, "O31", "تذکر", bold=True, bg=C_WARN)
    cell(ws, "O32:O33", "تعهد", bold=True, color="FFFFFF", bg=C_PLEDGE)
    cell(ws, "P31:S31", None, size=11, field="disc_7_warning")
    cell(ws, "P32:S33", None, size=11, field="disc_7_pledge")

    # ── 7. جمع بندی وپیشنهادات مدیریتی (U–X) ──────────────────────────────
    cell(ws, "U23:X23", "جمع بندی وپیشنهادات مدیریتی", bold=True, size=13,
         color="FF0000", bg=C_HEAD)
    # (label, symbol, symbol colour) – symbols stand in for the template's drawn shapes.
    recommendations = [
        ("شایسته تقدیروتشویق", "✸", "FF0000"),
        ("مناسب استمراردرمسئولیت فعلی", "★", "F4B000"),
        ("شایسته ارتقای شغلی", "⚑", "E0C000"),
        ("مستعدپذیرش مسئولیت بالاتر", "☁", "4E9A2E"),
        ("نیازمندبرنامه بهبودعملکرد", "■", "2E75B6"),
        ("نیازمندبررسی مجددمسئولیت", "■", "2E75B6"),
    ]
    for i, (label, sym, col) in enumerate(recommendations, 1):
        r = 23 + i
        cell(ws, f"U{r}:V{r}", label, size=11)
        cell(ws, f"W{r}", sym, size=16, color=col, font=SYMBOL_FONT)
        cell(ws, f"X{r}", "☐", size=14, field=f"recommend_{i}_mark", font=SYMBOL_FONT)
    cell(ws, "U30:X30", "توضیحات تکمیلی:", bold=True, h="right")
    cell(ws, "U31:X33", None, size=11, bg=C_PEACH, h="right", v="top",
         field="recommend_notes")

    # ── 8. سوابق ارزیابی ۵ سال گذشته (B–I) ───────────────────────────────
    cell(ws, "B35:I35", "سوابق ارزیابی ۵ سال گذشته", bold=True, size=13, bg=C_HEAD)
    for rng, t in (("B36", "سال"), ("C36", "نمره کل"), ("D36", "رتبه"),
                   ("E36:F36", "روندعملکرد"), ("G36:H36", "نقاط قوت شاخص"),
                   ("I36", "مواردنیازمند بهبود")):
        cell(ws, rng, t, bold=True, size=11, bg=C_SUB)
    for i, year in enumerate(("۱۴۰۰", "۱۴۰۱", "۱۴۰۲", "۱۴۰۳", "۱۴۰۴"), 1):
        r = 36 + i
        ws.row_dimensions[r].height = 32
        cell(ws, f"B{r}", year, bold=True)
        cell(ws, f"C{r}", None, field=f"hist_{i}_score")
        cell(ws, f"D{r}", None, field=f"hist_{i}_rank")
        cell(ws, f"E{r}:F{r}", None, size=10, field=f"hist_{i}_trend")
        cell(ws, f"G{r}:H{r}", None, size=10, field=f"hist_{i}_strengths")
        cell(ws, f"I{r}", None, size=10, field=f"hist_{i}_improve")
    cell(ws, "B42:D42", "میانگین ۵ ساله", bold=True, bg=C_SUB)
    cell(ws, "E42:I42", None, bold=True, field="hist_average")

    # ── 9. رضایت سنجی (K–S) ──────────────────────────────────────────────
    cell(ws, "K35:S35", "رضایت سنجی", bold=True, size=13, bg=C_HEAD)
    for rng, t in (("K36", "ردیف"), ("L36:N36", "شاخص"), ("O36", "امتیاز"),
                   ("P36:S36", "شرح")):
        cell(ws, rng, t, bold=True, bg=C_SUB)
    for i, label in enumerate(("رضایت مدیرمستقیم", "رضایت ارباب رجوع"), 1):
        r = 36 + i
        cell(ws, f"K{r}", _fa(i))
        cell(ws, f"L{r}:N{r}", label)
        cell(ws, f"O{r}", None, field=f"sat_{i}_score")
        cell(ws, f"P{r}:S{r}", None, size=11, field=f"sat_{i}_desc")
    cell(ws, "K39:S40",
         "منبع ارزیابی:سامانه سامع/ دفترارتباطات مردمی/مراجعات حضوری/تقدیرهاوشکایات",
         bold=True, size=10, h="right")

    # ── 10. جمع بندی نهائی ارزیابی عملکرد (U–X) ───────────────────────────
    cell(ws, "U35:X35", "جمع بندی نهائی ارزیابی عملکرد", bold=True, size=13,
         color="FF0000", bg=C_HEAD)
    final_items = [
        "امتیازعملکردواحد/شعبه", "امتیازعملکردفردی", "امتیازتوانمندی حرفه ای",
        "امتیازنظم وانضباط اداری", "امتیازرضایت مدیر", "امتیازتکریم ارباب رجوع",
    ]
    for i, label in enumerate(final_items, 1):
        r = 35 + i
        cell(ws, f"U{r}:W{r}", label, size=11)
        cell(ws, f"X{r}", None, bg=C_SCORE, field=f"final_{i}_score")

    # ── 11. نمره نهایی و امضا ────────────────────────────────────────────
    for r in range(44, 49):
        ws.row_dimensions[r].height = 26
    cell(ws, "B44:F48", "نمره نهایی ارزیابی عملکرد", bold=True, size=14, bg=C_HEAD)
    cell(ws, "G44:I48", None, bold=True, size=16, bg=C_SCORE, field="final_score")
    cell(ws, "K44:S44", "نام، سمت و امضای ارزیاب", bold=True, bg=C_SUB)
    cell(ws, "K45:S48", None, v="top", field="evaluator_signature")
    cell(ws, "U44:X44", "تاریخ", bold=True, bg=C_SUB)
    cell(ws, "U45:X48", None, field="evaluation_date")

    # ── Print setup ──────────────────────────────────────────────────────
    ws.print_area = "A1:Y49"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.row_breaks.append(Break(id=21))  # page 1: personal info, page 2: evaluation

    for out in OUT_DIRS:
        out.mkdir(exist_ok=True)
        wb.save(out / "report_template.xlsx")
        (out / "report_template_map.json").write_text(
            json.dumps({"sheet": SHEET, "fields": field_map, "images": image_boxes}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    print(f"wrote report_template.xlsx ({len(field_map)} fillable cells)")


def _fa(n: int) -> str:
    return str(n).translate(str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹"))


if __name__ == "__main__":
    build()
