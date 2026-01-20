"""Settings load/save helpers."""
from __future__ import annotations

import json
import logging
import re
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any, Dict

from app_paths import get_config_path, get_default_export_dir

logger = logging.getLogger(__name__)

def normalize_hex_color(value: Any) -> str | None:
    if not isinstance(value, str):
        return None

    raw = value.strip()
    if not raw:
        return None

    if raw.startswith("#"):
        raw = raw[1:]

    if len(raw) == 3:
        raw = "".join(ch * 2 for ch in raw)
    elif len(raw) != 6:
        return None

    if not re.fullmatch(r"[0-9a-fA-F]{6}", raw):
        return None

    return f"#{raw.upper()}"


def default_settings() -> Dict[str, Any]:
    export_folder = str(get_default_export_dir())
    return {
        "appearance_mode": "Dark",
        "theme_profile": "dark",
        "export_folder": export_folder,
        "themes": {
            "dark": {
                "colors": {
                    "background": "#1f1f1f",
                    "surface": "#2a2a2a",
                    "widget_fg": "#202225",
                    "text": "#ffffff",
                    "accent": "#1f6aa5",
                    "danger": "#8b0000",
                    "border": "#3a3a3a",
                    "scrollbar_track": "#2A2A2A",
                    "scrollbar_thumb": "#3A3A3A",
                    "scrollbar_thumb_hover": "#4A4A4A",
                    "header_bg": "#1E1E1E",
                    "header_text": "#FFFFFF",
                    "header_border": "#3A3A3A",
                    "selection_bg": "#1F6AA5",
                    "selection_text": "#FFFFFF",
                    "caret": "#FFFFFF",
                },
                "fonts": {
                    "family": "Segoe UI",
                    "base_size": 12,
                    "heading_size": 14,
                },
            },
            "light": {
                "colors": {
                    "background": "#f2f2f2",
                    "surface": "#ffffff",
                    "widget_fg": "#ffffff",
                    "text": "#111111",
                    "accent": "#1f6aa5",
                    "danger": "#b00020",
                    "border": "#d0d0d0",
                    "scrollbar_track": "#E6E6E6",
                    "scrollbar_thumb": "#C0C0C0",
                    "scrollbar_thumb_hover": "#AFAFAF",
                    "header_bg": "#FFFFFF",
                    "header_text": "#111111",
                    "header_border": "#D0D0D0",
                    "selection_bg": "#1F6AA5",
                    "selection_text": "#FFFFFF",
                    "caret": "#111111",
                },
                "fonts": {
                    "family": "Segoe UI",
                    "base_size": 12,
                    "heading_size": 14,
                },
            },
        },
    }


def _timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _backup_bad_file(path: Path) -> None:
    backup = path.with_name(f"{path.stem}.bad_{_timestamp()}{path.suffix}")
    try:
        path.replace(backup)
    except Exception:
        logger.exception("Не вдалося створити backup пошкодженого settings.json")


def _normalize_color(value: Any, fallback: str) -> str:
    normalized = normalize_hex_color(value)
    if normalized:
        return normalized
    return fallback


def _normalize_int(value: Any, fallback: int) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return fallback


def _merge_settings(base: Dict[str, Any], override: Dict[str, Any]) -> Dict[str, Any]:
    result = deepcopy(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _merge_settings(result[key], value)
        else:
            result[key] = value
    return result


def validate_settings(settings: Dict[str, Any]) -> Dict[str, Any]:
    defaults = default_settings()
    merged = _merge_settings(defaults, settings or {})

    appearance = merged.get("appearance_mode")
    if not isinstance(appearance, str) or appearance.title() not in {"System", "Light", "Dark"}:
        merged["appearance_mode"] = defaults["appearance_mode"]
    else:
        merged["appearance_mode"] = appearance.title()

    profile = merged.get("theme_profile")
    if not isinstance(profile, str) or profile.lower() not in {"dark", "light"}:
        merged["theme_profile"] = defaults["theme_profile"]
    else:
        merged["theme_profile"] = profile.lower()

    themes = merged.get("themes")
    if not isinstance(themes, dict):
        themes = {}
        merged["themes"] = themes

    for profile_key in ("dark", "light"):
        theme = themes.get(profile_key)
        if not isinstance(theme, dict):
            theme = {}
            themes[profile_key] = theme

        colors = theme.get("colors")
        if not isinstance(colors, dict):
            colors = {}
            theme["colors"] = colors

        default_colors = defaults["themes"][profile_key]["colors"]
        for color_key, fallback in default_colors.items():
            colors[color_key] = _normalize_color(colors.get(color_key), fallback)

        fonts = theme.get("fonts")
        if not isinstance(fonts, dict):
            fonts = {}
            theme["fonts"] = fonts

        default_fonts = defaults["themes"][profile_key]["fonts"]
        family = fonts.get("family")
        if not isinstance(family, str) or not family.strip():
            fonts["family"] = default_fonts["family"]
        else:
            fonts["family"] = family.strip()
        fonts["base_size"] = _normalize_int(fonts.get("base_size"), default_fonts["base_size"])
        fonts["heading_size"] = _normalize_int(fonts.get("heading_size"), default_fonts["heading_size"])

    export_folder = merged.get("export_folder")
    if not isinstance(export_folder, str) or not export_folder.strip():
        merged["export_folder"] = defaults["export_folder"]
    else:
        merged["export_folder"] = export_folder.strip()

    return merged


def load_settings() -> Dict[str, Any]:
    path = get_config_path("settings.json")
    defaults = default_settings()
    if not path.exists():
        settings = defaults
        try:
            save_settings(settings)
        except Exception:
            logger.exception("Не вдалося записати settings.json за замовчуванням")
        return settings

    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(raw, dict):
            raise ValueError("settings.json має містити об'єкт")
        return validate_settings(raw)
    except Exception:
        logger.exception("Не вдалося прочитати settings.json, створюємо дефолт")
        _backup_bad_file(path)
        settings = defaults
        try:
            save_settings(settings)
        except Exception:
            logger.exception("Не вдалося записати settings.json за замовчуванням")
        return settings


def save_settings(settings: Dict[str, Any]) -> None:
    path = get_config_path("settings.json")
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = validate_settings(settings)
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    tmp_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp_path.replace(path)
