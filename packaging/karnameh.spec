# PyInstaller spec — builds a one-folder, windowed app: dist/Karnameh/Karnameh.exe
# Build:  pyinstaller packaging/karnameh.spec --noconfirm
from pathlib import Path

ROOT = Path(SPECPATH).parent
PKG = ROOT / "karnameh"

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
    version=str(ROOT / "packaging" / "version_info.txt"),
)
coll = COLLECT(exe, a.binaries, a.datas, name="Karnameh", upx=False)
