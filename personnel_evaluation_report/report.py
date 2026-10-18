"""Report specification and the structured report model.

The single source of truth for the report: section order, official labels, which fields
come from the source workbook (auto), which are typed by the user (manual) and which are
static text. The edit form, the HTML view, the Markdown file and the Excel export are all
generated from `SECTIONS` + `FIELDS` + a validated `Report`.

No scores are calculated anywhere: totals, averages and the final score are manual
(confirmed by the user, 2026-10-16).
"""
from __future__ import annotations

import datetime as dt
import re
import uuid
from dataclasses import dataclass, field
from typing import Literal, Union

from . import jalali
from .normalize import clean_text, to_ascii_digits, to_fa_digits
from .photo import Photo
from .source import Employee, LoadedWorkbook

Kind = Literal["auto", "manual"]
Input = Literal["text", "textarea", "number", "count", "date", "check", "marital"]


@dataclass(frozen=True)
class FieldDef:
    id: str
    label: str
    kind: Kind
    input: Input = "text"
    source: tuple[str, ...] = ()   # source field ids (auto fields); joined with a space


FIELDS: dict[str, FieldDef] = {}


def _f(id: str, label: str, kind: Kind, input: Input = "text", source: tuple[str, ...] = ()) -> str:
    assert id not in FIELDS, id
    FIELDS[id] = FieldDef(id, label, kind, input, source)
    return id


# ── Field definitions (ids match Template/report_template_map.json) ───────
_f("full_name", "نام ونام خانوادگی", "auto", source=("first_name", "last_name"))
_f("national_id", "کدملی", "auto", source=("national_id",))
_f("personnel_no", "کدپرسنلی", "auto", source=("personnel_no",))
_f("birth_date", "تاریخ تولد", "auto", source=("birth_date",))
_f("degree", "مدرک تحصیلی", "auto", source=("degree",))
_f("field_of_study", "رشته تحصیلی", "auto", source=("field_of_study",))
_f("mobile", "شماره تماس", "auto", source=("mobile",))
_f("marital", "وضعیت تاهل", "auto", "marital", source=("marital_status",))
_f("children_daughters", "تعدادفرزندان — دختر", "manual", "count")
_f("children_sons", "تعدادفرزندان — پسر", "manual", "count")

_f("employment_type", "نوع استخدام", "auto", source=("employment_type",))
_f("post_title", "پست سازمانی", "auto", source=("post_title",))
_f("work_location", "محل خدمت", "auto", source=("work_dept",))
_f("start_date", "تاریخ شروع به کار", "auto", source=("divan_entry_date",))
_f("service_history", "سوابق خدمتی/تجربی", "manual", "textarea")
_f("training_history", "سوابق آموزشی", "manual", "textarea")
_f("eval_punishments", "سوابق ارزیابی — تنبیهات", "manual", "textarea")
_f("eval_rewards", "سوابق ارزیابی — تشویق ها", "manual", "textarea")
_f("previous_posts", "سمت های قبلی", "manual", "textarea")
_f("post_suspension", "توقف پست", "manual", "textarea")

COMPETENCIES = (
    "تخصص وتسلط بروظایف", "دقت وکیفیت انجام کار", "مسئولیت پذیری",
    "توانائی حل مسئله ومدیریت امور", "همکاری وکارتیمی", "تعامل وارتباط مؤثر",
    "ابتکاروارائه پیشنهادموثر",
)
for _i, _lbl in enumerate(COMPETENCIES, 1):
    _f(f"comp_{_i}_score", f"{_lbl} — نمره", "manual", "number")
    _f(f"comp_{_i}_desc", f"{_lbl} — شرح", "manual")
_f("comp_total", "جمع امتیاز توانمندی وشایستگی", "manual", "number")

DISCIPLINE = (
    "ورود وخروج به موقع", "میزان تأخیر/غیبت غیرموجه", "رعایت شئونات اداری",
    "پایبندی به اصول اخلاق حرفه ای", "آراستگی وپوشش", "شرکت درمراسم مذهبی_فرهنگی",
)
for _i, _lbl in enumerate(DISCIPLINE, 1):
    _f(f"disc_{_i}_score", f"{_lbl} — نمره", "manual", "number")
    _f(f"disc_{_i}_desc", f"{_lbl} — توضیحات/مستندات", "manual")
_f("disc_7_warning", "تعهدات وتذکرات — تذکر", "manual")
_f("disc_7_pledge", "تعهدات وتذکرات — تعهد", "manual")

RECOMMENDATIONS = (
    "شایسته تقدیروتشویق", "مناسب استمراردرمسئولیت فعلی", "شایسته ارتقای شغلی",
    "مستعدپذیرش مسئولیت بالاتر", "نیازمندبرنامه بهبودعملکرد", "نیازمندبررسی مجددمسئولیت",
)
for _i, _lbl in enumerate(RECOMMENDATIONS, 1):
    _f(f"recommend_{_i}_mark", _lbl, "manual", "check")
_f("recommend_notes", "توضیحات تکمیلی", "manual", "textarea")

HISTORY_YEARS = ("۱۴۰۰", "۱۴۰۱", "۱۴۰۲", "۱۴۰۳", "۱۴۰۴")   # fixed for now (user, 2026-10-16)
for _i, _y in enumerate(HISTORY_YEARS, 1):
    _f(f"hist_{_i}_score", f"سال {_y} — نمره کل", "manual", "number")
    _f(f"hist_{_i}_rank", f"سال {_y} — رتبه", "manual")
    _f(f"hist_{_i}_trend", f"سال {_y} — روندعملکرد", "manual")
    _f(f"hist_{_i}_strengths", f"سال {_y} — نقاط قوت شاخص", "manual")
    _f(f"hist_{_i}_improve", f"سال {_y} — مواردنیازمند بهبود", "manual")
_f("hist_average", "میانگین ۵ ساله", "manual", "number")

SATISFACTION = ("رضایت مدیرمستقیم", "رضایت ارباب رجوع")
for _i, _lbl in enumerate(SATISFACTION, 1):
    _f(f"sat_{_i}_score", f"{_lbl} — امتیاز", "manual", "number")
    _f(f"sat_{_i}_desc", f"{_lbl} — شرح", "manual")
SATISFACTION_NOTE = "منبع ارزیابی:سامانه سامع/ دفترارتباطات مردمی/مراجعات حضوری/تقدیرهاوشکایات"

FINAL_ITEMS = (
    "امتیازعملکردواحد/شعبه", "امتیازعملکردفردی", "امتیازتوانمندی حرفه ای",
    "امتیازنظم وانضباط اداری", "امتیازرضایت مدیر", "امتیازتکریم ارباب رجوع",
)
for _i, _lbl in enumerate(FINAL_ITEMS, 1):
    _f(f"final_{_i}_score", _lbl, "manual", "number")

_f("final_score", "نمره نهایی ارزیابی عملکرد", "manual", "number")
_f("evaluator_signature", "نام، سمت و امضای ارزیاب", "manual")
_f("evaluation_date", "تاریخ", "manual", "date")


# ── Section layout (official order) ──────────────────────────────────────

@dataclass(frozen=True)
class Static:
    text: str


Cell = Union[str, Static]   # str = field id


@dataclass(frozen=True)
class FieldsSection:
    title: str
    fields: tuple[str, ...]
    kind: str = "fields"
    layout: str = "kv"          # "kv": label/value rows; "row": labels as columns (signature block)
    photo: bool = False         # show the employee photo box beside the table


@dataclass(frozen=True)
class TableSection:
    title: str
    columns: tuple[str, ...]
    rows: tuple[tuple[Cell, ...], ...]
    note: str = ""
    kind: str = "table"


@dataclass(frozen=True)
class Banner:
    title: str
    kind: str = "banner"


def _fa(n: int) -> str:
    return to_fa_digits(str(n))


SECTIONS: tuple[Union[FieldsSection, TableSection, Banner], ...] = (
    FieldsSection("اطلاعات فردی", (
        "full_name", "national_id", "personnel_no", "birth_date", "degree",
        "field_of_study", "mobile", "marital", "children_daughters", "children_sons"), photo=True),
    FieldsSection("اطلاعات پرسنلی", (
        "employment_type", "post_title", "work_location", "start_date", "service_history",
        "training_history", "eval_punishments", "eval_rewards", "previous_posts", "post_suspension")),
    Banner("عملکردوتوانمندی کارکنان"),
    TableSection(
        "ارزیابی توانمندی وشایستگی", ("ردیف", "شاخص", "نمره", "شرح"),
        tuple((Static(_fa(i)), Static(lbl), f"comp_{i}_score", f"comp_{i}_desc")
              for i, lbl in enumerate(COMPETENCIES, 1))
        + ((Static(""), Static("جمع امتیاز"), "comp_total", Static("")),)),
    TableSection(
        "کارنامه نظم وانضباط اداری", ("ردیف", "شاخص", "نمره", "توضیحات/مستندات"),
        tuple((Static(_fa(i)), Static(lbl), f"disc_{i}_score", f"disc_{i}_desc")
              for i, lbl in enumerate(DISCIPLINE, 1))
        + ((Static(_fa(7)), Static("تعهدات وتذکرات"), Static("تذکر"), "disc_7_warning"),
           (Static(""), Static(""), Static("تعهد"), "disc_7_pledge"))),
    TableSection(
        "جمع بندی وپیشنهادات مدیریتی", ("گزینه", "انتخاب"),
        tuple((Static(lbl), f"recommend_{i}_mark") for i, lbl in enumerate(RECOMMENDATIONS, 1))
        + ((Static("توضیحات تکمیلی"), "recommend_notes"),)),
    TableSection(
        "سوابق ارزیابی ۵ سال گذشته",
        ("سال", "نمره کل", "رتبه", "روندعملکرد", "نقاط قوت شاخص", "مواردنیازمند بهبود"),
        tuple((Static(y), f"hist_{i}_score", f"hist_{i}_rank", f"hist_{i}_trend",
               f"hist_{i}_strengths", f"hist_{i}_improve") for i, y in enumerate(HISTORY_YEARS, 1))
        + ((Static("میانگین ۵ ساله"), "hist_average", Static(""), Static(""), Static(""), Static("")),)),
    TableSection(
        "رضایت سنجی", ("ردیف", "شاخص", "امتیاز", "شرح"),
        tuple((Static(_fa(i)), Static(lbl), f"sat_{i}_score", f"sat_{i}_desc")
              for i, lbl in enumerate(SATISFACTION, 1)),
        note=SATISFACTION_NOTE),
    TableSection(
        "جمع بندی نهائی ارزیابی عملکرد", ("شاخص", "امتیاز"),
        tuple((Static(lbl), f"final_{i}_score") for i, lbl in enumerate(FINAL_ITEMS, 1))),
    FieldsSection("نمره نهایی و امضا", ("final_score", "evaluator_signature", "evaluation_date"), layout="row"),
)

MARITAL_MARRIED = "married"
MARITAL_SINGLE = "single"
_MARITAL_SOURCE = {"متاهل": MARITAL_MARRIED, "متأهل": MARITAL_MARRIED, "مجرد": MARITAL_SINGLE}
MARITAL_LABELS = {MARITAL_MARRIED: "متأهل", MARITAL_SINGLE: "مجرد"}


# ── Report model ─────────────────────────────────────────────────────────

@dataclass
class Report:
    id: str
    snapshot_id: str
    source_file: str
    source_sheet: str
    source_row: int
    created_at: dt.datetime
    auto: dict[str, str]                       # auto field id → text from source ('' = missing)
    marital_source: str                        # raw source text
    marital: str                               # 'married' | 'single' | ''
    marital_from_source: bool
    source_issues: list[str]
    manual: dict[str, str] = field(default_factory=dict)   # cleaned user input
    errors: dict[str, str] = field(default_factory=dict)   # field id → Persian message
    validated: bool = False
    photo: Photo | None = None                              # employee photo (memory only)
    styles: dict = field(default_factory=dict)              # table appearance (see styles.py)

    # values used by every output
    def text(self, fid: str) -> str:
        """Display text (Persian digits for numbers the user typed)."""
        f = FIELDS[fid]
        if f.kind == "auto":
            if f.input == "marital":
                return MARITAL_LABELS.get(self.marital, self.marital_source)
            return self.auto.get(fid, "")
        v = self.manual.get(fid, "")
        if f.input in ("number", "count"):
            return fa_number(v)
        if f.input == "date":
            return to_fa_digits(v)
        if f.input == "check":
            return "☑" if v else "☐"
        return v

    def number(self, fid: str) -> float | int | None:
        v = self.manual.get(fid, "")
        if not v or FIELDS[fid].input not in ("number", "count"):
            return None
        n = float(v)
        return int(n) if n.is_integer() else n

    @property
    def created_fa(self) -> str:
        t = self.created_at
        return f"{jalali.from_gregorian(t.date()).format()} ساعت {to_fa_digits(t.strftime('%H:%M'))}"

    @property
    def employee_name(self) -> str:
        return self.auto.get("full_name", "")

    def warnings(self) -> list[str]:
        out = list(self.source_issues)
        empty_auto = [FIELDS[f].label for f, v in self.auto.items() if not v]
        if empty_auto:
            out.append("این موارد در فایل منبع خالی است: " + "، ".join(empty_auto))
        if not self.marital_from_source:
            if self.marital_source:
                out.append(f"وضعیت تاهل در فایل منبع «{self.marital_source}» است که با گزینه‌های فرم (متأهل/مجرد) مطابقت ندارد؛ "
                           "در صورت لزوم آن را دستی انتخاب کنید.")
            else:
                out.append("وضعیت تاهل در فایل منبع خالی است.")
        return out

    def empty_manual(self) -> list[str]:
        return [f.label for f in FIELDS.values()
                if f.kind == "manual" and f.input != "check" and not self.manual.get(f.id)]


def new_report(wb: LoadedWorkbook, emp: Employee) -> Report:
    auto: dict[str, str] = {}
    for f in FIELDS.values():
        if f.kind == "auto" and f.input != "marital":
            auto[f.id] = " ".join(p for p in (emp.values.get(s, "") for s in f.source) if p)
    marital_raw = emp.values.get("marital_status", "")
    marital = _MARITAL_SOURCE.get(clean_text(marital_raw), "")
    return Report(
        id=uuid.uuid4().hex,
        snapshot_id=wb.snapshot_id,
        source_file=wb.file_name,
        source_sheet=wb.sheet_name,
        source_row=emp.excel_row,
        created_at=dt.datetime.now(),
        auto=auto,
        marital_source=marital_raw,
        marital=marital,
        marital_from_source=bool(marital),
        source_issues=list(emp.issues),
    )


def fa_number(v: str) -> str:
    """Persian digits; decimal point shown as «/» (e.g. ۱۷/۵) as on Iranian score sheets.

    «٫» is avoided because in B Nazanin/B Mitra it is easily confused with the Persian zero «۰».
    """
    return to_fa_digits(v).replace(".", "/")


_NUMBER_RE = re.compile(r"^-?\d+(\.\d+)?$")


def _norm_number(s: str) -> str:
    s = to_ascii_digits(clean_text(s)).replace("٫", ".").replace("/", ".").replace(",", ".")
    return s.replace(" ", "")


def apply_form(report: Report, form: dict[str, str] | object) -> bool:
    """Replace all manual values from the submitted form and validate. Returns True if valid."""
    get = form.get  # works for dict and werkzeug MultiDict
    manual: dict[str, str] = {}
    errors: dict[str, str] = {}
    for f in FIELDS.values():
        if f.kind != "manual":
            continue
        raw = get(f.id) or ""
        if f.input == "check":
            manual[f.id] = "1" if raw else ""
            continue
        if f.input == "textarea":
            text = "\n".join(clean_text(line) for line in str(raw).replace("\r", "").split("\n")).strip()
        else:
            text = clean_text(raw)
        if not text:
            manual[f.id] = ""
            continue
        if f.input == "number":
            n = _norm_number(text)
            if not _NUMBER_RE.match(n):
                errors[f.id] = "فقط عدد وارد کنید."
            text = n
        elif f.input == "count":
            n = _norm_number(text)
            if not n.isdigit():
                errors[f.id] = "یک عدد صحیح صفر یا بزرگ‌تر وارد کنید."
            text = n
        elif f.input == "date":
            d = jalali.parse(text)
            if d is None:
                errors[f.id] = "تاریخ را به شکل سال/ماه/روز شمسی وارد کنید (مثلاً ۱۴۰۵/۰۷/۲۴)."
            else:
                text = to_ascii_digits(d.format())
        manual[f.id] = text

    if not report.marital_from_source:
        choice = get("marital") or ""
        report.marital = choice if choice in MARITAL_LABELS else ""

    report.manual = manual
    report.errors = errors
    report.validated = not errors
    return report.validated


def all_cells(section: TableSection) -> list[str]:
    return [c for row in section.rows for c in row if isinstance(c, str)]
