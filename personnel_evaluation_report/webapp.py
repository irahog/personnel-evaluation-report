"""Flask application: routes, per-browser-session state, security checks."""
from __future__ import annotations

import logging
import os
import secrets
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable
from urllib.parse import quote

from flask import (Flask, Response, abort, flash, g, redirect, render_template, request,
                   session, url_for)

from . import APP_TITLE_FA, VERSION, export_excel, export_md, filedialog, fonts, jalali, photo, styles, updater
from .normalize import normalize_national_id, to_fa_digits
from .paths import POSTS_MAPPING, RESOURCES
from .report import (FIELDS, MARITAL_LABELS, SECTIONS, Banner, FieldsSection, PhotoSection, Report, Static,
                     prefill_from_posts,
                     TableSection, apply_form, new_report)
from .settings import Settings
from .source import LoadedWorkbook, NeedsSheetChoice, SourceError, load_mapping, load_source

log = logging.getLogger("personnel_evaluation_report")


@dataclass
class SessionState:
    workbook: LoadedWorkbook | None = None
    sheet_choice: tuple[str, list[str]] | None = None   # (path, candidate sheet names)
    report: Report | None = None
    pending_replace: str = ""                             # sheet name awaiting confirmation
    posts: LoadedWorkbook | None = None                   # optional «فهرست کارکنان با پست قبلی»
    posts_sheet_choice: tuple[str, list[str]] | None = None


@dataclass
class AppState:
    settings: Settings
    sessions: dict[str, SessionState] = field(default_factory=dict)
    lock: threading.Lock = field(default_factory=threading.Lock)
    update_info: object = None                      # updater.ReleaseInfo from the last check
    update_job: object = None                       # updater.Job (background download)


def create_app(settings: Settings | None = None, *, port: int = 0, shutdown: Callable[[], None] | None = None,
               dialog: Callable[[str, str], str | None] | None = None) -> Flask:
    app = Flask(__name__, template_folder=str(RESOURCES / "templates"),
                static_folder=str(RESOURCES / "static"))
    app.secret_key = secrets.token_bytes(32)          # new every run: old cookies become invalid
    app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Strict",
                      MAX_CONTENT_LENGTH=photo.MAX_UPLOAD_BYTES + 1024 * 1024)   # photo + form fields
    app.jinja_env.tests["instanceof"] = isinstance
    app.jinja_env.filters["font_slug"] = fonts.slug
    app.jinja_env.filters["font_class"] = lambda fam: "f-" + fonts.slug(fam)
    app.jinja_env.filters["fa_time"] = lambda t: f"{jalali.from_gregorian(t.date()).format()} ساعت {to_fa_digits(t.strftime('%H:%M'))}"
    state = AppState(settings or Settings.load())
    ask = dialog or filedialog.ask
    app.extensions["personnel_evaluation_report"] = state

    # ── helpers ──────────────────────────────────────────────────────────
    def sess() -> SessionState:
        sid = session.get("sid")
        with state.lock:
            if not sid or sid not in state.sessions:
                sid = secrets.token_hex(16)
                session["sid"] = sid
                state.sessions[sid] = SessionState()
            return state.sessions[sid]

    def stale_message(s: SessionState) -> str:
        if s.workbook and s.workbook.is_stale():
            return ("فایل اطلاعات پرسنلی پس از بارگذاری تغییر کرده یا جابه‌جا شده است. "
                    "برای ادامه، دکمهٔ «بارگذاری دوباره» را بزنید.")
        return ""

    def posts_stale_message(s: SessionState) -> str:
        if s.posts and s.posts.is_stale():
            return ("فایل پست‌های قبلی پس از بارگذاری تغییر کرده یا جابه‌جا شده است. "
                    "دکمهٔ «بارگذاری دوباره» همان بخش را بزنید.")
        return ""

    def load_posts(s: SessionState, path: str, sheet: str | None = None) -> None:
        s.posts_sheet_choice = None
        try:
            wb = load_source(path, sheet, mapping=load_mapping(POSTS_MAPPING))
        except NeedsSheetChoice as e:
            s.posts_sheet_choice = (path, e.candidates)
            flash("فایل پست‌های قبلی: " + str(e), "warning")
            return
        except SourceError as e:
            flash("فایل پست‌های قبلی: " + str(e), "error")
            log.info("posts workbook load failed: %s", type(e).__name__)
            return
        s.posts = wb
        log.info("posts workbook loaded: rows=%d duplicates=%d", len(wb.employees), wb.duplicate_count)
        flash(f"فایل پست‌های قبلی «{wb.file_name}» با {to_fa_digits(str(len(wb.employees)))} ردیف بارگذاری شد.", "success")
        if s.report:
            prefill_from_posts(s.report, wb)
            s.report.validated = False
            flash("«سمت های قبلی» و «توقف پست» گزارش جاری از این فایل تکمیل شد (اگر خالی بودند).", "success")

    def require_report(s: SessionState, validated: bool = False) -> Report:
        if not s.workbook or not s.report:
            abort(redirect(url_for("index")))
        if s.report.snapshot_id != s.workbook.snapshot_id:
            s.report = None
            flash("گزارش جاری با فایل بارگذاری‌شده هم‌خوانی نداشت و پاک شد.", "error")
            abort(redirect(url_for("index")))
        if validated and not s.report.validated:
            flash("ابتدا فرم را بررسی و ثبت کنید.", "error")
            abort(redirect(url_for("edit_report")))
        return s.report

    def block_if_stale(s: SessionState, target: str = "index") -> None:
        msg = stale_message(s)
        if msg:
            flash(msg, "error")
            abort(redirect(url_for(target)))

    def load_into(s: SessionState, path: str, sheet: str | None = None) -> None:
        old_report = s.report
        s.sheet_choice = None
        try:
            wb = load_source(path, sheet)
        except NeedsSheetChoice as e:
            s.sheet_choice = (path, e.candidates)
            flash(str(e), "warning")
            return
        except SourceError as e:
            flash(str(e), "error")
            log.info("workbook load failed: %s", type(e).__name__)
            return
        s.workbook = wb
        s.report = None
        state.settings.last_dir = str(Path(path).parent)
        state.settings.save()
        log.info("workbook loaded: rows=%d duplicates=%d warnings=%d", len(wb.employees),
                 wb.duplicate_count, len(wb.warnings))
        flash(f"فایل «{wb.file_name}» با {to_fa_digits(str(len(wb.employees)))} ردیف کارمند بارگذاری شد.", "success")
        # Keep an open report only if the same employee is still uniquely present; refresh its source values.
        if old_report:
            matches = wb.find(old_report.auto.get("national_id", ""))
            if len(matches) == 1:
                fresh = new_report(wb, matches[0])
                fresh.manual = old_report.manual
                fresh.photo = old_report.photo
                fresh.styles = old_report.styles
                fresh.prefilled = old_report.prefilled
                fresh.posts_notes = old_report.posts_notes
                if not fresh.marital_from_source and not old_report.marital_from_source:
                    fresh.marital = old_report.marital
                s.report = fresh
                flash("اطلاعات گزارش جاری از فایل تازه به‌روزرسانی شد. لطفاً فرم را دوباره بررسی و ثبت کنید.", "warning")
            else:
                flash("کارمند گزارش جاری در فایل تازه به‌طور یکتا پیدا نشد و گزارش پاک شد.", "error")

    # ── request guards ───────────────────────────────────────────────────
    @app.before_request
    def guards():
        allowed = {f"127.0.0.1:{port}", f"localhost:{port}"} if port else None
        if allowed is not None and request.host not in allowed:
            abort(403)  # DNS-rebinding protection: only the loopback address we bound to
        if "csrf" not in session:
            session["csrf"] = secrets.token_urlsafe(32)
        if request.method == "POST" and request.form.get("_csrf") != session["csrf"]:
            abort(400)
        g.s = sess()

    @app.after_request
    def headers(resp: Response):
        resp.headers["Cache-Control"] = "no-store"
        resp.headers["X-Frame-Options"] = "DENY"
        resp.headers["X-Content-Type-Options"] = "nosniff"
        resp.headers["Referrer-Policy"] = "no-referrer"
        resp.headers["Content-Security-Policy"] = (
            "default-src 'self'; style-src 'self'; script-src 'self'; img-src 'self' data: blob:; "
            "font-src 'self'; form-action 'self'; frame-ancestors 'none'")
        return resp

    @app.context_processor
    def inject():
        s: SessionState = g.get("s") or SessionState()
        return {
            "csrf": session.get("csrf", ""), "s": s, "settings": state.settings, "fa": to_fa_digits,
            "title": APP_TITLE_FA, "version": VERSION, "stale": stale_message(s),
            "font_css": quote(state.settings.font), "posts_stale": posts_stale_message(s),
            "SECTIONS": SECTIONS, "FIELDS": FIELDS, "Static": Static, "Banner": Banner,
            "FieldsSection": FieldsSection, "TableSection": TableSection, "PhotoSection": PhotoSection,
            "MARITAL_LABELS": MARITAL_LABELS,
            "ST": styles,
        }

    @app.errorhandler(400)
    def bad_request(_e):
        return render_template("error.html", message="درخواست نامعتبر بود. صفحه را دوباره باز کنید و دوباره تلاش کنید."), 400

    @app.errorhandler(413)
    def too_large(_e):
        return render_template("error.html", message="حجم فایل ارسالی بیش از حد مجاز است. عکسی کوچک‌تر از ۱۵ مگابایت انتخاب کنید."), 413

    @app.errorhandler(403)
    def forbidden(_e):
        return "Forbidden", 403

    @app.errorhandler(Exception)
    def internal(e):
        from werkzeug.exceptions import HTTPException
        if isinstance(e, HTTPException):
            return e
        # Log type and stack frames only — never the exception message (it may contain cell values).
        tb = e.__traceback__
        frames = []
        while tb:
            frames.append(f"{Path(tb.tb_frame.f_code.co_filename).name}:{tb.tb_lineno}")
            tb = tb.tb_next
        log.error("internal error %s at %s", type(e).__name__, " > ".join(frames))
        return render_template("error.html", message="خطای غیرمنتظره‌ای رخ داد. عملیات انجام نشد. "
                                                     "برنامه را ببندید و دوباره باز کنید."), 500

    # ── home / source ────────────────────────────────────────────────────
    @app.get("/")
    def index():
        return render_template("index.html")

    @app.get("/health")
    def health():
        return {"app": "personnel-evaluation-report", "version": VERSION}

    @app.post("/source/pick")
    def pick_source():
        path = ask("open_source", state.settings.last_dir)
        if path:
            load_into(g.s, path)
        return redirect(url_for("index"))

    @app.post("/source/reload")
    def reload_source():
        s = g.s
        if s.workbook:
            load_into(s, s.workbook.path, s.workbook.sheet_name)
        return redirect(url_for("index"))

    @app.post("/source/sheet")
    def choose_sheet():
        s = g.s
        name = request.form.get("sheet", "")
        if s.sheet_choice and name in s.sheet_choice[1]:
            load_into(s, s.sheet_choice[0], name)
        return redirect(url_for("index"))

    # ── optional previous-posts workbook ────────────────────────────────
    @app.post("/posts/pick")
    def pick_posts():
        path = ask("open_posts", state.settings.last_dir)
        if path:
            if g.s.workbook and _same_file(path, g.s.workbook.path):
                flash("این همان فایل اطلاعات پرسنلی است؛ فایل «فهرست کارکنان با پست قبلی» را انتخاب کنید.", "error")
            else:
                load_posts(g.s, path)
        return redirect(url_for("index"))

    @app.post("/posts/reload")
    def reload_posts():
        if g.s.posts:
            load_posts(g.s, g.s.posts.path, g.s.posts.sheet_name)
        return redirect(url_for("index"))

    @app.post("/posts/sheet")
    def choose_posts_sheet():
        s = g.s
        name = request.form.get("sheet", "")
        if s.posts_sheet_choice and name in s.posts_sheet_choice[1]:
            load_posts(s, s.posts_sheet_choice[0], name)
        return redirect(url_for("index"))

    @app.post("/posts/clear")
    def clear_posts():
        g.s.posts = None
        g.s.posts_sheet_choice = None
        flash("فایل پست‌های قبلی کنار گذاشته شد؛ «سمت های قبلی» و «توقف پست» دستی پر می‌شوند.", "success")
        return redirect(url_for("index"))

    # ── search ───────────────────────────────────────────────────────────
    @app.post("/search")
    def search():
        s = g.s
        if not s.workbook:
            flash("ابتدا فایل اطلاعات پرسنلی را انتخاب کنید.", "error")
            return redirect(url_for("index"))
        block_if_stale(s)
        if posts_stale_message(s):
            flash(posts_stale_message(s), "error")
            return redirect(url_for("index"))
        raw = request.form.get("national_id", "")
        key = normalize_national_id(raw)
        if not key:
            flash("شماره ملی را وارد کنید.", "error")
            return redirect(url_for("index"))
        if not key.isdigit() or len(key) != 10:
            flash("شماره ملی باید ۱۰ رقم باشد و فقط شامل عدد باشد.", "error")
            return redirect(url_for("index"))
        matches = s.workbook.find(key)
        if not matches:
            flash("کارمندی با این شماره ملی در فایل پیدا نشد.", "error")
            return redirect(url_for("index"))
        if len(matches) > 1:
            rows = "، ".join(to_fa_digits(str(e.excel_row)) for e in matches)
            flash(f"این شماره ملی در فایل تکراری است (ردیف‌های اکسل: {rows}). "
                  "گزارش ساخته نشد؛ لطفاً فایل منبع را اصلاح و دوباره بارگذاری کنید.", "error")
            return redirect(url_for("index"))
        s.report = new_report(s.workbook, matches[0])   # replaces any previous report entirely
        if s.posts:
            prefill_from_posts(s.report, s.posts)
        s.pending_replace = ""
        log.info("report started")
        return redirect(url_for("edit_report"))

    # ── report ───────────────────────────────────────────────────────────
    @app.route("/report/edit", methods=["GET", "POST"])
    def edit_report():
        s = g.s
        r = require_report(s)
        if request.method == "POST":
            block_if_stale(s, "edit_report")
            ok = apply_form(r, request.form)
            r.styles, style_errors = styles.parse_form(request.form)
            for msg in style_errors:
                flash(msg, "warning")
            if request.form.get("photo_remove"):
                r.photo = None
            upload = request.files.get("photo")
            if upload and upload.filename:
                try:
                    r.photo = photo.process_upload(upload.read())
                except photo.PhotoError as e:
                    r.errors["photo"] = str(e)
                    r.validated = ok = False
            if ok:
                return redirect(url_for("view_report"))
            flash("برخی مقادیر نادرست است. موارد مشخص‌شده را اصلاح کنید.", "error")
            return render_template("edit.html", r=r), 422
        return render_template("edit.html", r=r)

    @app.get("/report/photo")
    def report_photo():
        r = require_report(g.s)
        if not r.photo:
            abort(404)
        return Response(r.photo.jpeg, mimetype="image/jpeg")

    @app.get("/report/style.css")
    def report_style():
        r = require_report(g.s)
        return Response(styles.css(r.styles), mimetype="text/css")

    @app.get("/report/view")
    def view_report():
        s = g.s
        r = require_report(s, validated=True)
        return render_template("view.html", r=r)

    @app.get("/report/download.md")
    def download_md():
        s = g.s
        r = require_report(s, validated=True)
        block_if_stale(s, "view_report")
        body = export_md.render(r).encode("utf-8")
        log.info("markdown exported")
        return _download(body, f"کارنامه - {r.employee_name}.md", "text/markdown; charset=utf-8")

    @app.get("/report/download.xlsx")
    def download_xlsx():
        s = g.s
        r = require_report(s, validated=True)
        block_if_stale(s, "view_report")
        fd = export_excel.decide_font(state.settings.font, state.settings.fallback_font)
        try:
            body = export_excel.build_single_workbook(r, fd.used)
        except export_excel.ExportError as e:
            flash(str(e), "error")
            return redirect(url_for("view_report"))
        log.info("excel exported")
        if fd.warning:
            flash(fd.warning, "warning")
        return _download(body, f"کارنامه - {r.employee_name}.xlsx",
                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    @app.post("/report/shared")
    def append_shared():
        s = g.s
        r = require_report(s, validated=True)
        block_if_stale(s, "view_report")
        path = state.settings.shared_workbook
        if not path:
            flash("ابتدا در «تنظیمات» کارنامهٔ مشترک را انتخاب یا ایجاد کنید.", "error")
            return redirect(url_for("view_report"))
        replace = request.form.get("replace") == "1" and s.pending_replace != ""
        fd = export_excel.decide_font(state.settings.font, state.settings.fallback_font)
        try:
            res = export_excel.append_to_shared(path, r, fd.used, replace=replace)
        except export_excel.SameEmployeeExists as e:
            s.pending_replace = e.sheet
            flash(str(e) + " برای جایگزینی، دکمهٔ «جایگزینی برگهٔ قبلی» را بزنید.", "warning")
            return redirect(url_for("view_report"))
        except export_excel.ExportError as e:
            flash(str(e), "error")
            return redirect(url_for("view_report"))
        s.pending_replace = ""
        log.info("shared workbook updated replaced=%s", res.replaced)
        verb = "جایگزین شد" if res.replaced else "افزوده شد"
        flash(f"برگهٔ «{res.sheet}» به کارنامهٔ مشترک {verb}.", "success")
        if fd.warning:
            flash(fd.warning, "warning")
        return redirect(url_for("view_report"))

    @app.post("/report/new")
    def new():
        g.s.report = None
        g.s.pending_replace = ""
        flash("گزارش قبلی پاک شد. شماره ملی کارمند بعدی را وارد کنید.", "success")
        return redirect(url_for("index"))

    # ── settings ─────────────────────────────────────────────────────────
    @app.route("/settings", methods=["GET", "POST"])
    def settings_page():
        st = state.settings
        if request.method == "POST":
            font, fallback = request.form.get("font", ""), request.form.get("fallback_font", "")
            if font in fonts.FONTS and fallback in fonts.FONTS:
                st.font, st.fallback_font = font, fallback
                st.save()
                flash("تنظیمات ذخیره شد.", "success")
            return redirect(url_for("settings_page"))
        return render_template("settings.html", font_status=fonts.status(), fonts=fonts.FONTS,
                               update_info=state.update_info, update_job=state.update_job,
                               can_self_update=updater.can_self_update(), releases_page=updater.RELEASES_PAGE,
                               shared_exists=bool(st.shared_workbook) and os.path.exists(st.shared_workbook))

    @app.post("/settings/shared")
    def pick_shared():
        mode = "new_shared" if request.form.get("mode") == "new" else "open_shared"
        start = str(Path(state.settings.shared_workbook).parent) if state.settings.shared_workbook else state.settings.last_dir
        path = ask(mode, start)
        if path:
            if not path.lower().endswith(".xlsx"):
                flash("کارنامهٔ مشترک باید فایل ‎.xlsx‎ باشد.", "error")
            elif s_is_source(path):
                flash("فایل اطلاعات پرسنلی را نمی‌توان به‌عنوان کارنامهٔ مشترک انتخاب کرد.", "error")
            else:
                state.settings.shared_workbook = path
                state.settings.save()
                flash("کارنامهٔ مشترک تنظیم شد.", "success")
        return redirect(url_for("settings_page"))

    def s_is_source(path: str) -> bool:
        wb = g.s.workbook
        try:
            return bool(wb) and os.path.samefile(path, wb.path)
        except OSError:
            return False

    @app.post("/settings/fonts-folder")
    def open_fonts_folder():
        from .paths import FONTS_DIR
        try:
            fonts.open_in_file_manager(FONTS_DIR)
        except OSError:
            flash("باز کردن پوشهٔ فونت‌ها ممکن نشد.", "error")
        return redirect(url_for("settings_page"))

    @app.post("/settings/install-fonts")
    def install_fonts():
        try:
            done = fonts.install_user_fonts()
        except OSError:
            log.info("font install failed")
            flash("نصب فونت‌ها ممکن نشد. فایل‌های فونت را از «باز کردن پوشهٔ فونت‌های برنامه» دستی نصب کنید.", "error")
            return redirect(url_for("settings_page"))
        log.info("fonts installed count=%d", len(done))
        if done:
            flash(f"{to_fa_digits(str(len(done)))} فایل فونت برای کاربر فعلی نصب شد. اگر اکسل باز است، آن را ببندید و دوباره باز کنید.", "success")
        else:
            flash("فونت‌های همراه برنامه از قبل نصب بودند.", "success")
        return redirect(url_for("settings_page"))

    # ── in-app update (user-initiated only) ─────────────────────────────
    @app.post("/settings/update/check")
    def update_check():
        try:
            info = updater.check()
        except updater.UpdateError as e:
            flash(str(e), "error")
            return redirect(url_for("settings_page") + "#update")
        state.update_info = info
        log.info("update check: latest=%s newer=%s", info.version, info.newer)
        if info.newer:
            flash(f"نسخهٔ تازهٔ {to_fa_digits(info.version)} منتشر شده است.", "success")
        else:
            flash(f"برنامه به‌روز است (نسخهٔ {to_fa_digits(VERSION)}).", "success")
        return redirect(url_for("settings_page") + "#update")

    @app.post("/settings/update/download")
    def update_download():
        info = state.update_info
        if not (info and info.newer and updater.can_self_update()):
            return redirect(url_for("settings_page") + "#update")
        job = state.update_job
        if job is None or job.state in ("idle", "error"):
            job = state.update_job = updater.Job(state="downloading", version=info.version,
                                                 message="در حال دانلود نسخهٔ تازه…")
            threading.Thread(target=updater.download, args=(info, job), daemon=True).start()
        return redirect(url_for("settings_page") + "#update")

    @app.get("/settings/update/status")
    def update_status():
        job = state.update_job
        return job.as_dict() if job else {"state": "idle", "percent": 0, "message": "", "version": ""}

    @app.post("/settings/update/install")
    def update_install():
        job = state.update_job
        if not (job and job.state == "ready" and job.file):
            return redirect(url_for("settings_page") + "#update")
        try:
            cmd = updater.prepare_install(job.file)
            updater.launch(cmd)
        except updater.UpdateError as e:
            flash(str(e), "error")
            return redirect(url_for("settings_page") + "#update")
        except OSError:
            flash("اجرای نصب‌کننده ممکن نشد. نسخهٔ تازه را از صفحهٔ دانلود دستی نصب کنید.", "error")
            return redirect(url_for("settings_page") + "#update")
        job.state = "installing"
        with state.lock:
            state.sessions.clear()
        log.info("update install started version=%s", job.version)
        if shutdown:
            threading.Timer(1.0, shutdown).start()
        return render_template("bye.html", updating=job.version)

    # ── quit ─────────────────────────────────────────────────────────────
    @app.post("/quit")
    def quit_app():
        with state.lock:
            state.sessions.clear()
        log.info("shutdown requested")
        if shutdown:
            threading.Timer(0.5, shutdown).start()
        return render_template("bye.html")

    return app


def _same_file(a: str, b: str) -> bool:
    try:
        return os.path.samefile(a, b)
    except OSError:
        return False


def _download(body: bytes, filename: str, mimetype: str) -> Response:
    safe = filename.translate(str.maketrans({c: " " for c in '\\/:*?"<>|'}))
    resp = Response(body, mimetype=mimetype)
    resp.headers["Content-Disposition"] = f"attachment; filename=\"report{Path(safe).suffix}\"; filename*=UTF-8''{quote(safe)}"
    return resp
