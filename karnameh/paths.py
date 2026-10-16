"""Filesystem locations (works both from source and from a PyInstaller bundle)."""
from __future__ import annotations

import os
import sys
from pathlib import Path

from . import APP_NAME


def resource_dir() -> Path:
    """Directory containing the `karnameh` package resources (templates, static, config)."""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS")) / "karnameh"
    return Path(__file__).resolve().parent


def user_data_dir() -> Path:
    """Per-user folder for settings and logs only — never employee data."""
    if sys.platform == "win32":
        base = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_DATA_HOME") or Path.home() / ".local" / "share")
    path = base / APP_NAME
    path.mkdir(parents=True, exist_ok=True)
    return path


RESOURCES = resource_dir()
TEMPLATE_XLSX = RESOURCES / "resources" / "report_template.xlsx"
TEMPLATE_MAP = RESOURCES / "resources" / "report_template_map.json"
SOURCE_MAPPING = RESOURCES / "config" / "source_mapping.json"
FONTS_DIR = RESOURCES / "static" / "fonts"
SAMPLE_WORKBOOK = RESOURCES / "resources" / "sample_personnel.xlsx"
