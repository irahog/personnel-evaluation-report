# PyInstaller spec — one-folder, windowed app.
#   Windows: dist/Karnameh/Karnameh.exe
#   macOS:   dist/Karnameh.app  (menu-bar-less background app; the UI is the browser)
# Build:  pyinstaller packaging/karnameh.spec --noconfirm
import sys
from pathlib import Path

ROOT = Path(SPECPATH).parent
sys.path.insert(0, str(ROOT))
from karnameh import VERSION  # noqa: E402
PKG = ROOT / "karnameh"
IS_WIN = sys.platform == "win32"
IS_MAC = sys.platform == "darwin"

datas = [
    (str(PKG / "templates"), "karnameh/templates"),
    (str(PKG / "static"), "karnameh/static"),
    (str(PKG / "config"), "karnameh/config"),
    (str(PKG / "resources"), "karnameh/resources"),
]

a = Analysis(
    [str(ROOT / "run_karnameh.py")],
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
    name="Karnameh",
    console=False,          # no terminal window for the user
    upx=False,
    version=str(ROOT / "packaging" / "version_info.txt") if IS_WIN else None,
)
coll = COLLECT(exe, a.binaries, a.datas, name="Karnameh", upx=False)

if IS_MAC:
    app = BUNDLE(
        coll,
        name="Karnameh.app",
        bundle_identifier="local.karnameh",
        version=VERSION,
        info_plist={
            "CFBundleDisplayName": "Karnameh",
            "CFBundleShortVersionString": VERSION,
            "CFBundleVersion": VERSION,
            # The UI is the browser: no Dock icon or menu bar, so macOS never reports the
            # app as "not responding". Quit from the app's «خروج از برنامه» button.
            "LSUIElement": True,
            "NSHighResolutionCapable": True,
            "LSMinimumSystemVersion": "11.0",
        },
    )
