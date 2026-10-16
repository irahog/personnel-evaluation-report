"""Supported report fonts: bundled files for the browser, installation check for Excel."""
from __future__ import annotations

import os
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
