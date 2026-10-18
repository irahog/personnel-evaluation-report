"""Supported report fonts: bundled files for the browser, installation check for Excel."""
from __future__ import annotations

import os
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path

from .paths import FONTS_DIR


@dataclass(frozen=True)
class FontInfo:
    family: str              # exact family name (used in CSS and in Excel cell formatting)
    label: str               # Persian label for the settings screen
    regular: str | None      # bundled file name (None = not bundled, must be installed)
    bold: str | None
    system_files: tuple[str, ...]   # file names that indicate the font is installed


FONTS: dict[str, FontInfo] = {
    f.family: f
    for f in (
        FontInfo("B Nazanin", "بی نازنین (B Nazanin)", "BNazanin.ttf", "BNaznnBd.ttf",
                 ("BNazanin.ttf", "BNaznnBd.ttf", "BNAZANIN.TTF")),
        FontInfo("B Mitra", "بی میترا (B Mitra)", "BMITRA.TTF", "BMitraBd.ttf",
                 ("BMITRA.TTF", "BMitraBd.ttf", "BMitra.ttf")),
        FontInfo("Vazirmatn FD", "وزیرمتن با ارقام فارسی (Vazirmatn FD)", "Vazirmatn-FD-Regular.ttf",
                 "Vazirmatn-FD-Bold.ttf", ("Vazirmatn-FD-Regular.ttf",)),
        # B Titr is a display face: one (bold) weight, used for both regular and bold text.
        FontInfo("B Titr", "بی تیتر (B Titr)", "BTitrBd.ttf", "BTitrBd.ttf",
                 ("BTitrBd.ttf", "BTITRBD.TTF", "BTitr.ttf")),
        # IRANSansX (Persian-digit variant) is a commercial FontIran font: never bundled; used when
        # the organisation has installed its licensed copy.
        FontInfo("IRANSansXFaNum", "ایران‌سنس ایکس با ارقام فارسی (IRANSansX FaNum)", None, None,
                 ("IRANSansXFaNum-Regular.ttf", "IRANSansXFaNum-Regular.otf", "IRANSansXFaNum.ttf",
                  "IRANSansX-FaNum-Regular.ttf")),
        # Times New Roman ships with Windows and may not be redistributed, so it is not bundled.
        FontInfo("Times New Roman", "تایمز نیو رومن (Times New Roman)", None, None,
                 ("times.ttf", "Times New Roman.ttf")),
    )
}
DEFAULT_FONT = "B Nazanin"
DEFAULT_FALLBACK = "Times New Roman"
SYMBOL_FONT = "Segoe UI Symbol"   # ☐ ☑ ★ … are not in the Persian fonts


def slug(family: str) -> str:
    return family.lower().replace(" ", "-")


def font_dirs() -> list[Path]:
    if sys.platform == "win32":
        dirs = [Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"]
        local = os.environ.get("LOCALAPPDATA")
        if local:
            dirs.append(Path(local) / "Microsoft" / "Windows" / "Fonts")
        return dirs
    if sys.platform == "darwin":
        return [Path.home() / "Library" / "Fonts", Path("/Library/Fonts"),
                Path("/System/Library/Fonts/Supplemental")]
    return [Path.home() / ".fonts", Path.home() / ".local/share/fonts", Path("/usr/share/fonts")]


def _registry_has(family: str) -> bool:
    if sys.platform != "win32":
        return False
    import winreg  # noqa: PLC0415 – Windows only

    key_path = r"Software\Microsoft\Windows NT\CurrentVersion\Fonts"
    for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        try:
            with winreg.OpenKey(hive, key_path) as key:
                i = 0
                while True:
                    try:
                        name, _, _ = winreg.EnumValue(key, i)
                    except OSError:
                        break
                    if name.lower().startswith(family.lower() + " ") or name.lower() == family.lower():
                        return True
                    i += 1
        except OSError:
            continue
    return False


def is_installed(family: str) -> bool:
    """Whether Excel on this computer can render `family` (font installed for the OS)."""
    info = FONTS.get(family)
    if info is None:
        return False
    if _registry_has(family):
        return True
    names = {n.lower() for n in info.system_files}
    for d in font_dirs():
        try:
            if any(p.name.lower() in names for p in d.iterdir()):
                return True
        except OSError:
            continue
    return False


# Full font names as Windows registers them (value name in the Fonts registry key).
FULL_NAMES = {
    "BNazanin.ttf": "B Nazanin", "BNaznnBd.ttf": "B Nazanin Bold",
    "BMITRA.TTF": "B Mitra", "BMitraBd.ttf": "B Mitra Bold",
    "Vazirmatn-FD-Regular.ttf": "Vazirmatn FD Regular", "Vazirmatn-FD-Bold.ttf": "Vazirmatn FD Bold",
    "BTitrBd.ttf": "B Titr Bold",
}


def user_font_dir() -> Path:
    """Per-user font folder: no administrator rights needed."""
    if sys.platform == "win32":
        return Path(os.environ["LOCALAPPDATA"]) / "Microsoft" / "Windows" / "Fonts"
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Fonts"
    return Path.home() / ".local" / "share" / "fonts"


def install_user_fonts() -> list[str]:
    """Install the bundled fonts for the current user so Excel can use them.

    Already-present files are left alone. Returns the file names that were installed.
    Windows: copy to the per-user font folder and register under HKCU (Windows 10 1803+).
    macOS: copy to ~/Library/Fonts (what Font Book's "Install" does for one user).
    """
    target = user_font_dir()
    target.mkdir(parents=True, exist_ok=True)
    done = []
    for name, src in bundled_files().items():
        dst = target / name
        if not dst.exists():
            shutil.copyfile(src, dst)
            done.append(name)
        if sys.platform == "win32":
            _register_windows_font(FULL_NAMES[name], dst)
    if done and sys.platform == "win32":
        _broadcast_font_change()
    return done


def _register_windows_font(full_name: str, path: Path) -> None:
    import winreg  # noqa: PLC0415 – Windows only

    key_path = r"Software\Microsoft\Windows NT\CurrentVersion\Fonts"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
        winreg.SetValueEx(key, f"{full_name} (TrueType)", 0, winreg.REG_SZ, str(path))


def _broadcast_font_change() -> None:
    """Tell running programs that fonts changed (Excel may still need a restart)."""
    try:
        import ctypes
        HWND_BROADCAST, WM_FONTCHANGE, SMTO_ABORTIFHUNG = 0xFFFF, 0x001D, 0x0002
        ctypes.windll.user32.SendMessageTimeoutW(HWND_BROADCAST, WM_FONTCHANGE, 0, 0,
                                                SMTO_ABORTIFHUNG, 1000, None)
    except Exception:  # noqa: BLE001 – cosmetic only
        pass


def open_in_file_manager(path: Path) -> None:
    import subprocess
    if sys.platform == "win32":
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def bundled_files() -> dict[str, Path]:
    out = {}
    for f in FONTS.values():
        for name in (f.regular, f.bold):
            if name:
                out[name] = FONTS_DIR / name
    return out


def status() -> list[dict]:
    """For the settings page."""
    return [
        {"family": f.family, "label": f.label, "bundled": f.regular is not None,
         "installed": is_installed(f.family)}
        for f in FONTS.values()
    ]
