"""Synthetic fixtures only — no real personnel data anywhere in the test suite."""
from __future__ import annotations

import datetime as dt
from pathlib import Path

import pytest
from openpyxl import Workbook

from personnel_evaluation_report.normalize import to_fa_digits as fa

HEADERS = [
    "ردیف", "نام", "نام خانوادگی", "نوع استخدام", "وضعیت اشتغال", "شماره پرسنلی", "شماره ملی",
    "تاریخ تولد", "عنوان پست سازمانی در ابلاغ", "واحد سازمانی در ابلاغ", "تاریخ آخرین انتصاب",
    "مدرک تحصیلی", "رشته تحصیلی", "وضعیت تاهل", "تعداد اولاد در ابلاغ ریالی", "تعداد اولاد واقعی",
    "محل خدمت ـ اداره کل", "محل خدمت ـ اداره/بخش", "تاریخ ورود به دیوان", "روز خدمت در دیوان",
    "ماه خدمت در دیوان", "سال خدمت در دیوان", "تاریخ استخدام در خدمات دولتی", "روزسابقه خدمت دولتی",
    "ماه سابقه خدمت دولتی", "سال سابقه خدمت دولتی", "تلفن همراه",
]


def person(i: int, **over) -> dict:
    d = {
        "ردیف": str(i), "نام": f"نام{i}", "نام خانوادگی": f"خانوادگی{i}", "نوع استخدام": "رسمی",
        "وضعیت اشتغال": "شاغل", "شماره پرسنلی": fa(f"900{i:03d}"), "شماره ملی": fa(f"0000000{i:03d}"),
        "تاریخ تولد": "۱۳۶۵/۰۲/۱۵", "عنوان پست سازمانی در ابلاغ": "کارشناس آزمایشی",
        "واحد سازمانی در ابلاغ": "واحد آزمایشی", "تاریخ آخرین انتصاب": "۱۴۰۲/۰۱/۰۱",
        "مدرک تحصیلی": "کارشناسی", "رشته تحصیلی": "رشتهٔ آزمایشی", "وضعیت تاهل": "مجرد",
        "تعداد اولاد در ابلاغ ریالی": "-", "تعداد اولاد واقعی": "۰",
        "محل خدمت ـ اداره کل": f"اداره کل آزمایشی {i}", "محل خدمت ـ اداره/بخش": f"بخش آزمایشی {i}",
        "تاریخ ورود به دیوان": "۱۳۹۵/۰۷/۰۱", "روز خدمت در دیوان": "۱", "ماه خدمت در دیوان": "۲",
        "سال خدمت در دیوان": "۹", "تاریخ استخدام در خدمات دولتی": "۱۳۹۴/۰۱/۰۱",
        "روزسابقه خدمت دولتی": "۰", "ماه سابقه خدمت دولتی": "۶", "سال سابقه خدمت دولتی": "۱۰",
        "تلفن همراه": "۰۹۱۲۰۰۰۰۰۰۰",
    }
    d.update(over)
    return d


def write_workbook(path: Path, rows: list[dict], *, headers=HEADERS, title_rows: int = 1,
                   sheet: str = "Sheet1", extra_sheets: dict | None = None) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    for _ in range(title_rows):
        ws.append([])
    ws.append(list(headers))
    for r in rows:
        ws.append([r.get(h) for h in headers])
    for name, (hdrs, data) in (extra_sheets or {}).items():
        es = wb.create_sheet(name)
        es.append(list(hdrs))
        for r in data:
            es.append([r.get(h) for h in hdrs])
    wb.save(path)
    return path


@pytest.fixture
def make_wb(tmp_path):
    def _make(rows=None, **kw) -> Path:
        rows = rows if rows is not None else [person(i) for i in range(1, 6)]
        return write_workbook(tmp_path / kw.pop("name", "personnel.xlsx"), rows, **kw)
    return _make


@pytest.fixture
def settings(tmp_path):
    from personnel_evaluation_report.settings import Settings
    s = Settings.load(tmp_path / "settings.json")
    s.shared_workbook = str(tmp_path / "shared.xlsx")
    return s


class Client:
    """Flask test client wrapper that handles the CSRF token and the injected file dialog."""

    def __init__(self, app, answers: list):
        self.c = app.test_client()
        self.answers = answers
        self.c.get("/")
        with self.c.session_transaction() as s:
            self.token = s["csrf"]

    def post(self, url, data=None, **kw):
        return self.c.post(url, data={"_csrf": self.token, **(data or {})}, **kw)

    def get(self, url, **kw):
        return self.c.get(url, **kw)

    def load(self, path):
        self.answers.append(str(path))
        return self.post("/source/pick", follow_redirects=True)


@pytest.fixture
def client(settings):
    from personnel_evaluation_report.webapp import create_app
    answers: list = []
    app = create_app(settings, port=0, dialog=lambda mode, d: answers.pop(0) if answers else None)
    return Client(app, answers)


@pytest.fixture
def today():
    return dt.date(2026, 10, 16)
