# Personnel Evaluation Report — کارنامه پرسنلی

A local Persian (RTL) web app for Windows and macOS. It builds the «کارنامه پرسنلی» evaluation report for one
employee at a time from the organization's personnel Excel workbook. Source fields are filled
automatically, the user completes the manual fields in the browser, and the report can be:

- printed from the browser,
- downloaded as `.xlsx` (the official layout, one sheet),
- downloaded as `.md`,
- appended as a new sheet to a **shared workbook** (one sheet per employee).

PDF generation is out of scope by design: the browser's print dialog covers it.

- **Install & use on Windows (Persian):** [`README-Windows.md`](README-Windows.md)
- **Install & use on macOS (Persian):** [`README-macOS.md`](README-macOS.md)
- Full end-user guide (Persian): [`docs/fa/راهنمای کاربر.html`](docs/fa/راهنمای%20کاربر.html)
- Discovery notes and confirmed decisions: [`docs/STAGE1_DISCOVERY.md`](docs/STAGE1_DISCOVERY.md)

---

## Architecture

| Concern | Choice | Why |
|---|---|---|
| Server | Python 3.12, Flask 3.1, served by Waitress on `127.0.0.1` only | Small; full control of RTL HTML and print CSS; packages cleanly with PyInstaller |
| Excel I/O | openpyxl 3.1.5 | Reads `.xlsx` with cached formula values, writes styled templates, merges, print setup |
| UI | Server-rendered Jinja templates and ~40 lines of vanilla JS | No build step; works offline; CSP `default-src 'self'` |
| File picking | Native Tk dialog in a short-lived child process (`--dialog`) | The app needs the real **path** for staleness checks and for the shared workbook, which a browser `<input type=file>` cannot give |
| Packaging | PyInstaller one-folder app; Windows: Inno Setup per-user installer; macOS: `Personnel Evaluation Report.app` in a `.dmg` (arm64 and x86_64 builds) | No admin rights; fonts installed per user; no Python needed on the target computer |
| CI | GitHub Actions: `windows-2022`, `macos-14` (Apple Silicon), `macos-15-intel` | Tests on each OS, builds, smoke-tests the packaged app |

Rejected alternatives (see the discovery doc §12): FastAPI (async/OpenAPI not needed), Streamlit (weak RTL, print and
layout control), ASP.NET Core (fine but heavier for no gain in template filling), Node/Express (weaker
template-preserving Excel libraries).

### Data flow and lifecycle

```
user picks file ──► source.load_source(path)  ── reads bytes once ──► LoadedWorkbook (in memory, per browser session)
                                                                      │  sha256 snapshot id, mtime/size for staleness
search national ID ─► exact match on normalized key ─► report.new_report()  (replaces any previous report)
edit form ─► report.apply_form()  (validate, no calculations) ─► Report.validated
view / .md / .xlsx / shared ─► all rendered from the same Report + report.SECTIONS/FIELDS
```

- **No database and no cache.** Employee data lives only in the Flask process memory, keyed by a random
  per-browser session id. «خروج از برنامه» clears the sessions and exits the process. The secret key is new
  on every run, so old cookies are invalid.
- **Staleness.** Every search, form submission and export checks the source file's mtime and size. If the file
  has changed, the action is blocked until the user reloads. On reload, auto fields are refreshed from the new
  snapshot, manual values are kept, and the user must re-validate. A report carries its `snapshot_id`, so it can
  never mix workbook versions.
- **Temporary files.** Downloads are streamed from memory. The shared workbook is written to
  `~per-report-*.xlsx` next to the target and then `os.replace`d, so a crash never leaves it half-written. The
  temp file is removed on failure. The file-dialog child process passes the chosen path back through a temp file,
  which is deleted immediately.
- **On disk (`%LOCALAPPDATA%\PersonnelEvaluationReport`).** `settings.json` holds the font, fallback font, shared-workbook path and
  last folder. `logs/personnel-evaluation-report.log` holds counts, event names and exception type plus stack frames only (never the
  exception message, IDs, phones, dates or names). `instance.json` holds the port, for single-instance detection.

### Security

- Waitress binds `127.0.0.1` only. A Host-header check rejects anything except `127.0.0.1:<port>` and
  `localhost:<port>` (DNS-rebinding protection).
- Every POST requires a per-session CSRF token. Cookies are `HttpOnly` and `SameSite=Strict`.
- Responses send `Cache-Control: no-store`, a strict CSP, `X-Frame-Options: DENY`, `nosniff` and `no-referrer`.
- There are no external URLs anywhere in templates or static files, enforced by a test.

## Project layout

```
personnel_evaluation_report/
  __main__.py        launcher: single instance, port fallback, logging, --dialog, --selftest
  webapp.py          Flask routes, session state, guards, error handling
  source.py          workbook loading: sheet/header detection, cell conversion, duplicates, staleness
  report.py          THE report spec (SECTIONS, FIELDS) + Report model + form validation
  export_excel.py    fill template, single-sheet workbook, append to shared workbook
  export_md.py       Markdown renderer
  normalize.py       digit/letter/header normalization (explicit rules, no fuzzy matching)
  jalali.py          Jalali parse/validate, Gregorian→Jalali (33-year rule, 1178–1634 AP)
  fonts.py           supported fonts, Windows install detection (registry + font folders)
  filedialog.py      native dialog via child process
  settings.py        user preferences (no employee data)
  config/source_mapping.json   source column mapping (editable, see below)
  resources/         report_template.xlsx + report_template_map.json + sample_personnel.xlsx (synthetic)
  templates/, static/          UI, report view, print CSS, bundled fonts
tools/build_template.py        regenerates the blank report template and its cell map
packaging/                     PyInstaller spec, Inno Setup script, Farsi.isl, version info
tests/                         pytest suite (synthetic data only)
samples/                       synthetic sample outputs
docs/                          discovery notes, Persian user guide
```

## Source-to-report mapping

**Source columns** are configured in `personnel_evaluation_report/config/source_mapping.json`. For each field it lists the
**exact** accepted header spellings, the type and whether the field is required. Matching ignores only the
following explicit differences: surrounding/multiple spaces, Arabic «ي/ك» vs Persian «ی/ک», and dash variants
(`ـ – — -`), with spaces around dashes. The headers from the brief and the spellings found in the synthetic
workbook (`روز سابقه خدمت دولتی`, `ماه/سال سابقه خدمتی دولتی`, `-` instead of `ـ`) are both listed. Columns
are found by header, never by position. Extra columns are ignored. A missing required column stops loading,
with the column named in Persian.

**Report fields** are defined in `personnel_evaluation_report/report.py`:

| Report field | Source | Kind |
|---|---|---|
| نام ونام خانوادگی | نام + نام خانوادگی | auto |
| کدملی / کدپرسنلی / تاریخ تولد / مدرک تحصیلی / رشته تحصیلی | same-named source columns (شماره ملی, شماره پرسنلی, …) | auto |
| شماره تماس | تلفن همراه | auto |
| وضعیت تاهل (☐ متأهل ☐ مجرد) | وضعیت تاهل; other values such as «معیل» or empty → user chooses | auto, manual only if unmatched |
| نوع استخدام | نوع استخدام | auto |
| پست سازمانی | عنوان پست سازمانی در ابلاغ | auto |
| محل خدمت | محل خدمت ـ اداره/بخش | auto (confirmed) |
| تاریخ شروع به کار | تاریخ ورود به دیوان | auto (confirmed) |
| تعدادفرزندان دختر/پسر | — | manual (confirmed) |
| everything else (histories, scores, totals, averages, recommendations, final score, signature, date) | — | manual; **nothing is calculated** (confirmed) |

`tools/build_template.py` writes `report_template.xlsx` and `report_template_map.json` (field id → cell).
`tests/test_outputs.py::test_spec_matches_template_map` keeps the spec and the template in sync.

Source values are displayed exactly as stored (Persian digits included). Conversions are explicit and flagged
to the user as per-employee notes:

- A numeric national-ID cell is zero-padded to 10 digits.
- A numeric phone cell is shown as is, with a warning; it is not padded.
- A real Excel date cell is converted from Gregorian to Jalali.
- An invalid Jalali text date is shown unchanged, with a warning.
- A formula cell without a cached value is shown empty, with a warning.

## Development

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements-dev.txt
.venv/bin/python -m pytest -q
.venv/bin/python -m personnel_evaluation_report                 # opens the browser
.venv/bin/python -m personnel_evaluation_report --selftest out  # end-to-end check with the bundled synthetic workbook
```

For UI work without native dialogs (e.g. driving the browser automatically), set
`PER_REPORT_DEV_DIALOG_OPEN_SOURCE=/path/to.xlsx` (and `…_OPEN_SHARED` / `…_NEW_SHARED`). The dialog then
returns that path. This is a development aid only.

## macOS build (CI)

`.github/workflows/macos.yml` builds on Apple Silicon (`macos-14`) and Intel (`macos-15-intel`). For each, it:

1. runs the pytest suite;
2. builds `Personnel Evaluation Report.app` and checks the binary's architecture;
3. self-tests the packaged app from a path with spaces and Persian characters;
4. installs the fonts with `--install-fonts` and checks that they're detected;
5. starts the server and checks that it listens on 127.0.0.1 only;
6. builds `PersonnelEvaluationReport-<version>-macOS-<arch>.dmg`, mounts it and self-tests the app inside;
7. uploads the `.dmg` as `personnel-evaluation-report-macos-<arch>`.

The app is an `LSUIElement` agent: it has no Dock icon and the browser is its UI. It is ad-hoc signed by PyInstaller
but not notarized, so Gatekeeper asks once on first launch (see `README-macOS.md`). Notarizing needs an Apple
Developer account.

Local build on a Mac: `pip install -r requirements-dev.txt`, then
`pyinstaller packaging/personnel_evaluation_report.spec --noconfirm`, which produces `dist/Personnel Evaluation Report.app`.

## Windows build (CI)

`.github/workflows/windows.yml` runs on every push. It:

1. runs the full pytest suite on Windows;
2. builds `dist/PersonnelEvaluationReport/` with PyInstaller (windowed, no console);
3. runs `PersonnelEvaluationReport.exe --selftest`;
4. starts the packaged server with port 8765 deliberately occupied, then checks that it picked another port,
   answers `/health`, serves an RTL page, and **listens on 127.0.0.1 only** (`Get-NetTCPConnection`);
5. builds the Inno Setup installer, installs it silently **per user**, self-tests the installed exe, checks that
   B Nazanin, B Mitra and Vazirmatn FD are registered for the user, and runs again from a path with spaces and
   Persian characters;
6. uploads `PersonnelEvaluationReport-Setup-*.exe`, `PersonnelEvaluationReport-portable.zip`, test results and the self-test outputs.

Local build on Windows: `pip install -r requirements-dev.txt`, then
`pyinstaller packaging/personnel_evaluation_report.spec --noconfirm`, then `iscc packaging\installer.iss` (Inno Setup 6.5+).

## Dependencies

Runtime dependencies are pinned in `requirements.txt`: Flask 3.1.3, Waitress 3.0.2, openpyxl 3.1.5 and their
pinned transitive dependencies. The build adds pytest 8.4.2 and PyInstaller 6.22.3. There are no other
services: no database, Docker, WSL or web server.

Bundled fonts:

| Font | Files | License |
|---|---|---|
| B Nazanin, B Mitra | `BNazanin.ttf`, `BNaznnBd.ttf`, `BMITRA.TTF`, `BMitraBd.ttf` | Widely distributed freeware from Borna Rayaneh. **Redistribution terms are not formally documented; confirm internally.** |
| Vazirmatn FD v33.003 | `Vazirmatn-FD-*.ttf` | SIL OFL 1.1 (`Vazirmatn-OFL.txt`) |
| Times New Roman | — | Not bundled (Microsoft license); ships with Windows |

The ☐ ☑ ★ symbols use Windows' `Segoe UI Symbol`, because the Persian fonts do not contain them.

## Verification status

See the handover summary, which states what was run where. As of this commit:

- The pytest suite passes on macOS with Python 3.10.
- The browser UI was driven in a real browser.
- The print layout was checked through Chrome's print-to-PDF.
- The template was rendered by Microsoft Excel for Mac.
- The Windows build, installer and font installation are verified by the Windows CI workflow.
- The macOS `.app` was built and self-tested on this Apple Silicon Mac; the Intel build is verified only by CI.
- The native file dialog inside the packaged `.app` was checked to start without crashing, but its on-screen
  appearance has not been checked automatically.

## Known limitations

- Drawn shapes in the colleague's template (star, cloud, flag) are replaced by coloured Unicode symbols. Cell
  positions are reconstructed from photos.
- Shared workbooks containing shapes, images or form controls are refused, because openpyxl would drop them on save.
- The «جمع بندی نهائی ارزیابی عملکرد» list has the 6 rows visible in the photos. The years ۱۴۰۰–۱۴۰۴ are fixed.
- The Markdown output cannot reproduce merged cells, colours or symbols. It keeps all content and order.
- `.xlsx` cannot embed fonts. Excel needs the fonts installed, which the installer does for the current user.
- Partially completed reports are not persisted (by design).
- The executable is not code-signed, so Windows SmartScreen may warn on first run. The macOS app is not
  notarized, so Gatekeeper asks once on first launch.
- On macOS, Excel has no `Segoe UI Symbol`, so ☐ ☑ use the Mac's own symbol font.
