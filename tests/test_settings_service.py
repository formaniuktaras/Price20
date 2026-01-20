from __future__ import annotations

import json
from pathlib import Path

from settings_service import default_settings, load_settings, normalize_hex_color, save_settings, validate_settings


def test_load_settings_when_missing(tmp_path, monkeypatch):
    settings_path = tmp_path / "settings.json"
    monkeypatch.setattr("settings_service.get_config_path", lambda _: settings_path)

    settings = load_settings()

    assert settings_path.exists()
    assert settings["appearance_mode"] == default_settings()["appearance_mode"]


def test_load_settings_backs_up_invalid_json(tmp_path, monkeypatch):
    settings_path = tmp_path / "settings.json"
    settings_path.write_text("{bad json", encoding="utf-8")
    monkeypatch.setattr("settings_service.get_config_path", lambda _: settings_path)

    settings = load_settings()

    backups = list(Path(tmp_path).glob("settings.bad_*.json"))
    assert backups, "Expected backup file for invalid settings.json"
    assert settings["appearance_mode"] == default_settings()["appearance_mode"]
    assert settings_path.exists()


def test_save_settings_atomic(tmp_path, monkeypatch):
    settings_path = tmp_path / "settings.json"
    monkeypatch.setattr("settings_service.get_config_path", lambda _: settings_path)

    payload = default_settings()
    payload["appearance_mode"] = "Light"
    save_settings(payload)

    assert settings_path.exists()
    assert not settings_path.with_suffix(".json.tmp").exists()
    data = json.loads(settings_path.read_text(encoding="utf-8"))
    assert data["appearance_mode"] == "Light"


def test_normalize_hex_color():
    assert normalize_hex_color("1f6aa5") == "#1F6AA5"
    assert normalize_hex_color("#abc") == "#AABBCC"
    assert normalize_hex_color("zzz") is None


def test_validate_settings_tolerates_partial_theme():
    settings = {"themes": {"dark": {"colors": {"accent": "#abc"}}}}
    validated = validate_settings(settings)
    assert validated["themes"]["dark"]["colors"]["accent"] == "#AABBCC"
