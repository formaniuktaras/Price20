"""Application paths and logging setup helpers."""
from __future__ import annotations

import logging
import os
import platform
import shutil
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Iterable

APP_NAME = "ProdGen"
_DATA_ENV_VARS = ("PRODGEN_DATA_DIR", "PRICE16_DATA_DIR")
_DEBUG_ENV_VARS = ("PRODGEN_DEBUG", "PRICE16_DEBUG")


def _default_data_dir() -> Path:
    system = platform.system().lower()
    home = Path.home()
    if system == "windows":
        base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
        if base:
            return Path(base) / APP_NAME
        return home / APP_NAME
    if system == "darwin":
        return home / "Library" / "Application Support" / APP_NAME
    return home / ".local" / "share" / APP_NAME


def get_data_dir() -> Path:
    for var in _DATA_ENV_VARS:
        value = os.environ.get(var)
        if value:
            return Path(value).expanduser().resolve()
    return _default_data_dir()


def get_db_path() -> Path:
    return get_data_dir() / "catalog.db"


def get_config_path(filename: str) -> Path:
    return get_data_dir() / filename


def get_logs_dir() -> Path:
    return get_data_dir() / "logs"


def get_locks_dir() -> Path:
    return get_data_dir() / "locks"


def get_default_export_dir() -> Path:
    home = Path.home()
    documents = home / "Documents"
    downloads = home / "Downloads"
    if documents.exists():
        return documents
    if downloads.exists():
        return downloads
    return home


def setup_logging() -> None:
    logs_dir = get_logs_dir()
    logs_dir.mkdir(parents=True, exist_ok=True)
    log_path = logs_dir / "app.log"
    logger = logging.getLogger()
    if any(isinstance(handler, RotatingFileHandler) and handler.baseFilename == str(log_path) for handler in logger.handlers):
        return

    level = logging.DEBUG if any(os.environ.get(var) == "1" for var in _DEBUG_ENV_VARS) else logging.INFO
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
    handler = RotatingFileHandler(log_path, maxBytes=5 * 1024 * 1024, backupCount=5, encoding="utf-8")
    handler.setFormatter(formatter)

    logger.setLevel(level)
    logger.addHandler(handler)
    if not any(isinstance(h, logging.StreamHandler) for h in logger.handlers):
        stream_handler = logging.StreamHandler()
        stream_handler.setFormatter(formatter)
        stream_handler.setLevel(level)
        logger.addHandler(stream_handler)


LEGACY_FILES = (
    "catalog.db",
    "templates.json",
    "export_fields.json",
    "title_tags_templates.json",
)


def _timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _backup_path(path: Path) -> Path:
    return path.with_name(f"{path.name}.bak_{_timestamp()}")


def migrate_legacy_files(project_root: Path) -> None:
    data_dir = get_data_dir()
    data_dir.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger(__name__)

    for name in LEGACY_FILES:
        legacy_path = project_root / name
        if not legacy_path.exists():
            continue
        target = data_dir / name
        if target.exists():
            continue

        backup = _backup_path(target)
        try:
            shutil.copy2(legacy_path, target)
            shutil.copy2(legacy_path, backup)
            logger.info("Migrated legacy file %s to %s (backup at %s)", legacy_path, target, backup)
        except Exception:
            logger.exception("Failed to migrate legacy file %s", legacy_path)


__all__ = [
    "APP_NAME",
    "get_data_dir",
    "get_db_path",
    "get_config_path",
    "get_logs_dir",
    "get_default_export_dir",
    "get_locks_dir",
    "setup_logging",
    "migrate_legacy_files",
]
