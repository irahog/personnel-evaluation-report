"""End-to-end browser workflow, state isolation, security and startup."""
from __future__ import annotations

import json
import logging
import os
import socket
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

import pytest

from tests.conftest import person, write_workbook

ROOT = Path(__file__).resolve().parent.parent


def _search(client, nid):
    return client.post("/search", {"national_id": nid})


def test_full_workflow(client, make_wb, settings):
    r = client.load(make_wb())
    assert "بارگذاری شد" in r.get_data(as_text=True)
    r = _search(client, "۰۰۰۰۰۰۰۰۰۳")
    assert r.status_code == 302 and r.headers["Location"].endswith("/report/edit")
    page = client.get("/report/edit").get_data(as_text=True)
    assert "نام3 خانوادگی3" in page and 'dir="rtl"' in page and 'lang="fa"' in page
    r = client.post("/report/edit", {"comp_1_score": "۱۹", "final_score": "۹۰"})
    assert r.status_code == 302
    view = client.get("/report/view").get_data(as_text=True)
    assert "۱۹" in view and "بخش آزمایشی 3" in view
    assert client.get("/report/download.xlsx").data[:2] == b"PK"
    md = client.get("/report/download.md")
    assert md.mimetype == "text/markdown" and "۱۹" in md.get_data(as_text=True)
    assert "filename*=UTF-8''" in md.headers["Content-Disposition"]
    r = client.post("/report/shared", follow_redirects=True)
    assert "افزوده شد" in r.get_data(as_text=True) and Path(settings.shared_workbook).exists()


def test_view_requires_validated_form(client, make_wb):
    client.load(make_wb())
    _search(client, "0000000001")
    assert client.get("/report/view").headers["Location"].endswith("/report/edit")
    assert client.get("/report/download.xlsx").status_code == 302
    r = client.post("/report/edit", {"comp_1_score": "abc"})
    assert r.status_code == 422 and "فقط عدد" in r.get_data(as_text=True)
    assert client.get("/report/download.md").status_code == 302


@pytest.mark.parametrize("nid,msg", [("", "وارد کنید"), ("12345", "۱۰ رقم"), ("abcdefghij", "۱۰ رقم"),
                                     ("9999999999", "پیدا نشد")])
def test_search_messages(client, make_wb, nid, msg):
    client.load(make_wb())
    r = client.post("/search", {"national_id": nid}, follow_redirects=True)
    assert msg in r.get_data(as_text=True)


def test_duplicate_id_blocks_report(client, make_wb):
    client.load(make_wb([person(1), person(2, **{"شماره ملی": "0000000001"})]))
    r = client.post("/search", {"national_id": "0000000001"}, follow_redirects=True)
    body = r.get_data(as_text=True)
    assert "تکراری" in body and "ردیف‌های اکسل" in body
    assert client.get("/report/edit").status_code == 302


def test_search_without_workbook(client):
    r = client.post("/search", {"national_id": "0000000001"}, follow_redirects=True)
    assert "ابتدا فایل" in r.get_data(as_text=True)


def test_new_employee_does_not_inherit_previous_values(client, make_wb):
    client.load(make_wb())
    _search(client, "0000000001")
    client.post("/report/edit", {"comp_1_score": "۱۷", "recommend_notes": "یادداشت کارمند اول",
                                 "children_sons": "۲"})
    _search(client, "0000000002")
    page = client.get("/report/edit").get_data(as_text=True)
    assert "یادداشت کارمند اول" not in page and "۱۷" not in page and "نام1" not in page
    assert "نام2 خانوادگی2" in page
    client.post("/report/new")
    assert client.get("/report/edit").status_code == 302


def test_sessions_are_isolated(settings, make_wb):
    from karnameh.webapp import create_app
    from tests.conftest import Client
    answers: list = []
    app = create_app(settings, dialog=lambda m, d: answers.pop(0))
    a, b = Client(app, answers), Client(app, answers)
    a.load(make_wb())
    _search(a, "0000000001")
    assert b.get("/report/edit").status_code == 302          # b has no report
    assert "نام1" not in b.get("/").get_data(as_text=True)


def test_csrf_required(client, make_wb):
    r = client.c.post("/search", data={"national_id": "0000000001"})
    assert r.status_code == 400


def test_host_header_check(settings):
    from karnameh.webapp import create_app
    app = create_app(settings, port=8765)
    c = app.test_client()
    assert c.get("/", headers={"Host": "127.0.0.1:8765"}).status_code == 200
    assert c.get("/", headers={"Host": "evil.example:8765"}).status_code == 403


def test_security_headers_and_no_cache(client):
    r = client.get("/")
    assert r.headers["Cache-Control"] == "no-store"
    assert "default-src 'self'" in r.headers["Content-Security-Policy"]
    assert r.headers["X-Frame-Options"] == "DENY"


def test_stale_source_blocks_until_reload(client, make_wb):
    p = make_wb()
    client.load(p)
    _search(client, "0000000001")
    client.post("/report/edit", {"comp_1_score": "۱۵"})
    write_workbook(p, [person(1, **{"محل خدمت ـ اداره/بخش": "بخش تازه"}), person(2)])
    st = os.stat(p)
    os.utime(p, ns=(st.st_atime_ns, st.st_mtime_ns + 5_000_000_000))
    r = client.get("/report/download.xlsx", follow_redirects=True)
    assert "تغییر کرده" in r.get_data(as_text=True)
    r = client.post("/search", {"national_id": "0000000002"}, follow_redirects=True)
    assert "تغییر کرده" in r.get_data(as_text=True)
    # Controlled reload: source values refresh, manual input kept, must re-validate.
    r = client.post("/source/reload", follow_redirects=True)
    assert "به‌روزرسانی شد" in r.get_data(as_text=True)
    page = client.get("/report/edit").get_data(as_text=True)
    assert "بخش تازه" in page and "۱۵" in page
    assert client.get("/report/view").status_code == 302


def test_sheet_choice_flow(client, tmp_path):
    from tests.conftest import HEADERS
    p = write_workbook(tmp_path / "x.xlsx", [person(1)], sheet="الف", extra_sheets={"ب": (HEADERS, [person(2)])})
    r = client.load(p)
    assert "برگهٔ درست" in r.get_data(as_text=True)
    r = client.post("/source/sheet", {"sheet": "ب"}, follow_redirects=True)
    assert "بارگذاری شد" in r.get_data(as_text=True)


def test_bad_file_message(client, tmp_path):
    p = tmp_path / "bad.xlsx"
    p.write_bytes(b"nope")
    assert "خراب" in client.load(p).get_data(as_text=True)


def test_shared_not_configured(client, make_wb, settings):
    settings.shared_workbook = ""
    client.load(make_wb())
    _search(client, "0000000001")
    client.post("/report/edit", {})
    r = client.post("/report/shared", follow_redirects=True)
    assert "تنظیمات" in r.get_data(as_text=True)


def test_shared_replace_needs_confirmation(client, make_wb, settings):
    client.load(make_wb())
    _search(client, "0000000001")
    client.post("/report/edit", {})
    client.post("/report/shared")
    r = client.post("/report/shared", follow_redirects=True)
    assert "جایگزینی برگهٔ قبلی" in r.get_data(as_text=True)
    r = client.post("/report/shared", {"replace": "1"}, follow_redirects=True)
    assert "جایگزین شد" in r.get_data(as_text=True)


def test_export_failure_shows_no_success(client, make_wb, settings, monkeypatch):
    from karnameh import export_excel
    client.load(make_wb())
    _search(client, "0000000001")
    client.post("/report/edit", {})

    def boom(*a, **k):
        raise export_excel.ExportError("ذخیرهٔ کارنامهٔ مشترک ممکن نشد.")
    monkeypatch.setattr(export_excel, "append_to_shared", boom)
    body = client.post("/report/shared", follow_redirects=True).get_data(as_text=True)
    assert "ممکن نشد" in body and "افزوده شد" not in body


def test_settings_font_selection(client):
    r = client.post("/settings", {"font": "B Mitra", "fallback_font": "Times New Roman"}, follow_redirects=True)
    body = r.get_data(as_text=True)
    assert 'class="f-b-mitra fb-times-new-roman"' in body
    client.post("/settings", {"font": "Comic Sans", "fallback_font": "Times New Roman"})
    assert 'class="f-b-mitra' in client.get("/").get_data(as_text=True)


def test_unexpected_error_is_generic_and_logged_without_values(client, make_wb, monkeypatch, caplog):
    from karnameh import export_md
    client.load(make_wb())
    _search(client, "0000000001")
    client.post("/report/edit", {})

    def boom(r):
        raise ValueError("0000000001 sensitive")
    monkeypatch.setattr(export_md, "render", boom)
    with caplog.at_level(logging.INFO, logger="karnameh"):
        r = client.get("/report/download.md")
    assert r.status_code == 500 and "0000000001" not in r.get_data(as_text=True)
    assert "Traceback" not in r.get_data(as_text=True)
    assert "0000000001" not in caplog.text and "ValueError" in caplog.text


def test_logs_never_contain_personal_data(client, make_wb, caplog):
    with caplog.at_level(logging.DEBUG):
        client.load(make_wb())
        _search(client, "0000000004")
        client.post("/report/edit", {"comp_1_score": "۱۸"})
        client.get("/report/download.xlsx")
        client.get("/report/download.md")
        client.post("/report/shared")
    text = caplog.text
    for secret in ("0000000004", "۰۰۰۰۰۰۰۰۰۴", "نام4", "۰۹۱۲۰۰۰۰۰۰۰", "۱۳۶۵/۰۲/۱۵"):
        assert secret not in text


def test_print_css_present():
    css = (ROOT / "karnameh/static/app.css").read_text(encoding="utf-8")
    assert "@media print" in css and "@page" in css and ".no-print" in css


def test_quit_clears_sessions(settings, make_wb):
    from karnameh.webapp import create_app
    from tests.conftest import Client
    called = []
    answers: list = []
    app = create_app(settings, dialog=lambda m, d: answers.pop(0), shutdown=lambda: called.append(1))
    c = Client(app, answers)
    c.load(make_wb())
    assert "برنامه بسته شد" in c.post("/quit").get_data(as_text=True)
    assert app.extensions["karnameh"].sessions == {}
    time.sleep(0.8)
    assert called


def test_free_port_handles_conflict():
    from karnameh.__main__ import PREFERRED_PORTS, free_port
    held = []
    try:
        for p in PREFERRED_PORTS:
            s = socket.socket()
            try:
                s.bind(("127.0.0.1", p))
                held.append(s)
            except OSError:
                s.close()
        port = free_port()
        assert port not in PREFERRED_PORTS or not held
    finally:
        for s in held:
            s.close()


def _lan_ip() -> str | None:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(("10.255.255.255", 1))   # no packet is sent for UDP connect
            ip = s.getsockname()[0]
            return None if ip.startswith("127.") else ip
    except OSError:
        return None


def test_application_startup_binds_loopback_only(tmp_path):
    env = {**os.environ, "HOME": str(tmp_path), "LOCALAPPDATA": str(tmp_path), "XDG_DATA_HOME": str(tmp_path)}
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        port = s.getsockname()[1]
    proc = subprocess.Popen([sys.executable, "-m", "karnameh", "--no-browser", "--port", str(port)],
                            cwd=ROOT, env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        for _ in range(50):
            try:
                with urllib.request.urlopen(f"http://127.0.0.1:{port}/health", timeout=1) as r:
                    assert json.loads(r.read())["app"] == "karnameh"
                    break
            except OSError:
                time.sleep(0.2)
        else:
            pytest.fail("server did not start")
        ip = _lan_ip()
        if ip:   # the same server must not answer on the computer's network address
            opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
            with pytest.raises(OSError):
                opener.open(f"http://{ip}:{port}/health", timeout=2)
        # Second launch detects the running instance instead of starting another server.
        r = subprocess.run([sys.executable, "-m", "karnameh", "--no-browser"], cwd=ROOT, env=env, timeout=30)
        assert r.returncode == 0
    finally:
        proc.terminate()
        proc.wait(timeout=10)


def test_selftest_command(tmp_path):
    env = {**os.environ, "HOME": str(tmp_path), "LOCALAPPDATA": str(tmp_path), "XDG_DATA_HOME": str(tmp_path)}
    r = subprocess.run([sys.executable, "-m", "karnameh", "--selftest", str(tmp_path / "out")], cwd=ROOT, env=env,
                       capture_output=True, text=True, timeout=120)
    assert r.returncode == 0, r.stdout + r.stderr
    assert (tmp_path / "out" / "report.xlsx").exists()


def test_server_is_created_on_loopback(monkeypatch, tmp_path):
    import karnameh.__main__ as m
    import waitress
    seen = {}

    class Fake:
        def run(self):
            pass

    def fake_create_server(app, **kw):
        seen.update(kw)
        return Fake()
    monkeypatch.setattr(waitress, "create_server", fake_create_server)
    monkeypatch.setattr(m, "existing_instance", lambda: None)
    monkeypatch.setattr(m, "user_data_dir", lambda: tmp_path)
    m.serve(open_browser=False, port=8799)
    assert seen["host"] == "127.0.0.1" and seen["port"] == 8799


def test_no_external_resources_offline():
    """Pages must work offline: no CDN, web font, analytics or other external URL."""
    import re
    for p in list((ROOT / "karnameh/templates").glob("*.html")) + list((ROOT / "karnameh/static").glob("*.*")):
        text = p.read_text(encoding="utf-8")
        assert not re.search(r"(https?:)?//[a-z0-9.-]+\.[a-z]{2,}", text, re.I), p.name
