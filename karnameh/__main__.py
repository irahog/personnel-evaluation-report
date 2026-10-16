"""Program entry point (also the PyInstaller entry point).

  karnameh                     start the local server and open the browser
  karnameh --no-browser        start without opening a browser
  karnameh --dialog MODE DIR OUT   (internal) show a native file dialog
  karnameh --selftest OUTDIR   run an end-to-end check on the bundled synthetic workbook
"""
from __future__ import annotations

import argparse
import json
import logging
import logging.handlers
import os
import socket
import sys
import threading
import urllib.request
import webbrowser

from . import VERSION
from .paths import user_data_dir

PREFERRED_PORTS = range(8765, 8775)
log = logging.getLogger("karnameh")


def setup_logging() -> None:
    logdir = user_data_dir() / "logs"
    logdir.mkdir(exist_ok=True)
    h = logging.handlers.RotatingFileHandler(logdir / "karnameh.log", maxBytes=1_000_000, backupCount=3, encoding="utf-8")
    h.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    root.addHandler(h)
    # waitress/werkzeug access logs would contain URLs only (no form data), but keep them quiet anyway.
    logging.getLogger("waitress").setLevel(logging.WARNING)


def _instance_file():
    return user_data_dir() / "instance.json"


def existing_instance() -> int | None:
    """Port of an already-running instance, if any."""
    try:
        port = int(json.loads(_instance_file().read_text(encoding="utf-8"))["port"])
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/health", timeout=2) as r:
            if json.loads(r.read()).get("app") == "karnameh":
                return port
    except Exception:  # noqa: BLE001 – any failure means "no usable instance"
        return None
    return None


def free_port() -> int:
    for p in PREFERRED_PORTS:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", p))
                return p
            except OSError:
                continue
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:   # all busy: let the OS choose
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def serve(open_browser: bool, port: int | None = None) -> None:
    from waitress import create_server

    from .webapp import create_app

    running = existing_instance()
    if running:
        log.info("already running; opening browser")
        if open_browser:
            webbrowser.open(f"http://127.0.0.1:{running}/")
        return

    port = port or free_port()
    url = f"http://127.0.0.1:{port}/"

    def shutdown() -> None:
        try:
            _instance_file().unlink(missing_ok=True)
        finally:
            log.info("stopped")
            logging.shutdown()
            os._exit(0)   # process exit releases all in-memory employee data

    app = create_app(port=port, shutdown=shutdown)
    server = create_server(app, host="127.0.0.1", port=port, threads=4)   # loopback only
    _instance_file().write_text(json.dumps({"port": port, "pid": os.getpid()}), encoding="utf-8")
    log.info("started version=%s port=%d", VERSION, port)
    if open_browser:
        threading.Timer(0.8, webbrowser.open, args=(url,)).start()
    try:
        server.run()
    finally:
        _instance_file().unlink(missing_ok=True)


def selftest(outdir: str) -> int:
    """Exercise the real code paths (load → search → form → all outputs) with synthetic data."""
    from pathlib import Path

    from . import export_md
    from .paths import SAMPLE_WORKBOOK
    from .settings import Settings
    from .webapp import create_app

    out = Path(outdir)
    out.mkdir(parents=True, exist_ok=True)
    shared = out / "shared.xlsx"
    shared.unlink(missing_ok=True)
    (out / "selftest.txt").unlink(missing_ok=True)
    settings = Settings.load(out / "settings.json")
    settings.shared_workbook = str(shared)
    answers = iter([str(SAMPLE_WORKBOOK)])
    app = create_app(settings, dialog=lambda mode, d: next(answers))
    c = app.test_client()
    c.get("/")
    with c.session_transaction() as s:
        token = s["csrf"]
    steps = []

    def check(name: str, ok: bool) -> None:
        steps.append((name, ok))
        line = ("PASS " if ok else "FAIL ") + name
        print(line)
        with open(out / "selftest.txt", "a", encoding="utf-8") as fh:   # windowed .exe has no console
            fh.write(line + "\n")

    r = c.post("/source/pick", data={"_csrf": token}, follow_redirects=True)
    check("load sample workbook", "بارگذاری شد" in r.get_data(as_text=True))
    r = c.post("/search", data={"_csrf": token, "national_id": "۰۱۵۰۰۹۹۰۰۱"})
    check("search", r.status_code == 302 and r.headers["Location"].endswith("/report/edit"))
    r = c.post("/report/edit", data={"_csrf": token, "comp_1_score": "۱۸", "comp_total": "۱۸",
                                      "recommend_2_mark": "1", "final_score": "۸۵٫۵"})
    check("validate form", r.status_code == 302)
    check("edit form", c.get("/report/edit").status_code == 200)
    resp = c.get("/report/view")
    view = resp.get_data(as_text=True)
    check("html view", resp.status_code == 200 and 'class="report"' in view and "تخصص وتسلط بروظایف" in view)
    xlsx = c.get("/report/download.xlsx")
    (out / "report.xlsx").write_bytes(xlsx.data)
    check("excel export", xlsx.status_code == 200 and xlsx.data[:2] == b"PK")
    md = c.get("/report/download.md")
    (out / "report.md").write_bytes(md.data)
    check("markdown export", md.status_code == 200 and "# کارنامه پرسنلی" in md.get_data(as_text=True))
    (out / "report.html").write_text(view, encoding="utf-8")
    r = c.post("/report/shared", data={"_csrf": token}, follow_redirects=True)
    check("append to shared workbook", shared.exists() and "افزوده شد" in r.get_data(as_text=True))
    from .paths import FONTS_DIR, TEMPLATE_XLSX
    check("bundled resources", TEMPLATE_XLSX.exists() and len(list(FONTS_DIR.glob("*.[tT][tT][fF]"))) == 6)
    _ = export_md  # imported to make sure the module is bundled
    from . import fonts
    for f in fonts.status():
        print(f"font {f['family']}: installed={f['installed']}")
        with open(out / "selftest.txt", "a", encoding="utf-8") as fh:
            fh.write(f"INFO font {f['family']} installed={f['installed']}\n")
    return 0 if all(ok for _, ok in steps) else 1


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="karnameh")
    p.add_argument("--dialog", nargs=3, metavar=("MODE", "DIR", "OUT"))
    p.add_argument("--selftest", metavar="OUTDIR")
    p.add_argument("--no-browser", action="store_true")
    p.add_argument("--port", type=int)
    args = p.parse_args(argv)
    # A windowed (no-console) .exe has no stdout/stderr; give libraries somewhere harmless to write.
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")

    if args.dialog:
        from .filedialog import run_dialog
        run_dialog(*args.dialog)
        return 0
    setup_logging()
    if args.selftest:
        return selftest(args.selftest)
    serve(open_browser=not args.no_browser, port=args.port)
    return 0


if __name__ == "__main__":
    sys.exit(main())
