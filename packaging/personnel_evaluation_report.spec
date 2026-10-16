# PyInstaller spec — one-folder, windowed app.
#   Windows: dist/PersonnelEvaluationReport/PersonnelEvaluationReport.exe
#   macOS:   dist/Personnel Evaluation Report.app  (menu-bar-less background app; the UI is the browser)
# Build:  pyinstaller packaging/personnel_evaluation_report.spec --noconfirm
import sys
from pathlib import Path

ROOT = Path(SPECPATH).parent
sys.path.insert(0, str(ROOT))
from personnel_evaluation_report import VERSION  # noqa: E402
PKG = ROOT / "personnel_evaluation_report"
IS_WIN = sys.platform == "win32"
IS_MAC = sys.platform == "darwin"

datas = [
    (str(PKG / "templates"), "personnel_evaluation_report/templates"),
    (str(PKG / "static"), "personnel_evaluation_report/static"),
    (str(PKG / "config"), "personnel_evaluation_report/config"),
    (str(PKG / "resources"), "personnel_evaluation_report/resources"),
]

a = Analysis(
    [str(ROOT / "run_app.py")],
    pathex=[str(ROOT)],
    datas=datas,
    hiddenimports=["waitress", "tkinter", "tkinter.filedialog"],
    excludes=["pytest"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name="PersonnelEvaluationReport",
    console=False,          # no terminal window for the user
    upx=False,
    version=str(ROOT / "packaging" / "version_info.txt") if IS_WIN else None,
)
coll = COLLECT(exe, a.binaries, a.datas, name="PersonnelEvaluationReport", upx=False)

if IS_MAC:
    app = BUNDLE(
        coll,
        name="Personnel Evaluation Report.app",
        bundle_identifier="local.personnel-evaluation-report",
        version=VERSION,
        info_plist={
            "CFBundleDisplayName": "Personnel Evaluation Report",
            "CFBundleShortVersionString": VERSION,
            "CFBundleVersion": VERSION,
            # The UI is the browser: no Dock icon or menu bar, so macOS never reports the
            # app as "not responding". Quit from the app's «خروج از برنامه» button.
            "LSUIElement": True,
            "NSHighResolutionCapable": True,
            "LSMinimumSystemVersion": "11.0",
        },
    )
