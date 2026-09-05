"""Persistent user settings for the GUI."""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

from i18n import DEFAULT_LANGUAGE, normalize_language


def settings_path() -> Path:
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
    return base / "WordHighLightFormatInspect" / "settings.json"


def load_language(path: Path | None = None) -> str | None:
    target = path or settings_path()
    try:
        with target.open("r", encoding="utf-8") as file:
            value = json.load(file).get("language")
    except (OSError, json.JSONDecodeError, AttributeError):
        return None
    return value if isinstance(value, str) and value in ("de", "en", "fr") else None


def save_language(language: str, path: Path | None = None) -> None:
    target = path or settings_path()
    target.parent.mkdir(parents=True, exist_ok=True)
    with target.open("w", encoding="utf-8") as file:
        json.dump({"language": normalize_language(language)}, file, ensure_ascii=False, indent=2)
        file.write("\n")


def initial_language(path: Path | None = None, system_language: str = DEFAULT_LANGUAGE) -> str:
    return load_language(path) or normalize_language(system_language)