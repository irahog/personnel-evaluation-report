"""User preferences (font, shared-workbook path). Never contains employee data."""
from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path

from . import fonts
from .paths import user_data_dir


@dataclass
class Settings:
    font: str = fonts.DEFAULT_FONT
    fallback_font: str = fonts.DEFAULT_FALLBACK
    shared_workbook: str = ""
    last_dir: str = ""

    @classmethod
    def load(cls, path: Path | None = None) -> "Settings":
        path = path or user_data_dir() / "settings.json"
        s = cls()
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            for k in asdict(s):
                if isinstance(data.get(k), str):
                    setattr(s, k, data[k])
        except (OSError, ValueError):
            pass
        if s.font not in fonts.FONTS:
            s.font = fonts.DEFAULT_FONT
        if s.fallback_font not in fonts.FONTS:
            s.fallback_font = fonts.DEFAULT_FALLBACK
        s._path = path  # type: ignore[attr-defined]
        return s

    def save(self) -> None:
        path: Path = getattr(self, "_path", None) or user_data_dir() / "settings.json"
        path.write_text(json.dumps(asdict(self), ensure_ascii=False, indent=2), encoding="utf-8")
