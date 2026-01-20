"""Utilities for building the browser-based description editor."""
from __future__ import annotations

import argparse
import json
import logging
import os
import shutil
import subprocess
import sys
import threading
import time
from pathlib import Path
from typing import Iterable, Sequence

from app_paths import get_locks_dir

ROOT_DIR = Path(__file__).resolve().parent
EDITOR_ROOT = ROOT_DIR / "desc-editor"
DIST_DIR = EDITOR_ROOT / "dist"
NODE_MODULES = EDITOR_ROOT / "node_modules"
INSTALL_STAMP = NODE_MODULES / ".install-stamp"
BUILD_STAMP = DIST_DIR / ".build-stamp"
LOCK_MAX_AGE_SECONDS = 600
LOCK_FILENAME = "desc_editor_build.lock"

LOGGER = logging.getLogger(__name__)


class DescEditorBuildError(RuntimeError):
    """Raised when the description editor bundle cannot be produced."""


__all__ = ["ensure_desc_editor_built", "DescEditorBuildError"]

_BUILD_LOCK = threading.Lock()


def _iter_files(paths: Sequence[Path]) -> Iterable[Path]:
    for base in paths:
        if base.is_file():
            yield base
        elif base.is_dir():
            for path in base.rglob("*"):
                if path.is_file():
                    yield path


def _latest_mtime(paths: Sequence[Path]) -> float:
    latest = 0.0
    for path in _iter_files(paths):
        try:
            mtime = path.stat().st_mtime
        except FileNotFoundError:
            continue
        if mtime > latest:
            latest = mtime
    return latest


def _resolve_npm(explicit: str | None) -> str:
    candidate = explicit or "npm"
    resolved = shutil.which(candidate)
    if resolved is None:
        raise DescEditorBuildError(
            "Не знайдено npm. Встановіть Node.js з npm або передайте шлях через --npm."
        )
    return resolved


def _ensure_node_version() -> str:
    node_path = shutil.which("node")
    if node_path is None:
        raise DescEditorBuildError(
            "Не знайдено Node.js. Встановіть LTS-версію з https://nodejs.org (18 або новішу)."
        )
    try:
        completed = subprocess.run(
            [node_path, "--version"],
            check=True,
            capture_output=True,
            text=True,
        )
    except subprocess.CalledProcessError as exc:  # pragma: no cover - depends on external tools
        raise DescEditorBuildError("Не вдалося визначити версію Node.js (node --version завершився з помилкою).") from exc
    version = (completed.stdout or completed.stderr or "").strip()
    normalized = version.lstrip("vV")
    parts = normalized.split(".")
    try:
        major = int(parts[0])
    except (ValueError, IndexError):
        major = 0
    if major < 18:
        raise DescEditorBuildError(
            f"Потрібна версія Node.js 18 або новіша, знайдено {version or 'невідому версію'}."
        )
    return node_path


def _run_command(command: Sequence[str], *, cwd: Path, quiet: bool = False) -> None:
    stdout = subprocess.PIPE if quiet else None
    try:
        subprocess.run(command, cwd=str(cwd), check=True, stdout=stdout, stderr=subprocess.STDOUT if quiet else None)
    except subprocess.CalledProcessError as exc:  # pragma: no cover - depends on external tools
        output = exc.stdout.decode() if quiet and exc.stdout else ""
        raise DescEditorBuildError(f"Команда {' '.join(command)} завершилася з помилкою. {output}") from exc
    except FileNotFoundError as exc:
        raise DescEditorBuildError(f"Не вдалося виконати {' '.join(command)}: команда не знайдена") from exc


def _lock_file_path() -> Path:
    locks_dir = get_locks_dir()
    locks_dir.mkdir(parents=True, exist_ok=True)
    return locks_dir / LOCK_FILENAME


def _acquire_build_lock() -> Path:
    lock_path = _lock_file_path()
    now = time.time()
    while True:
        try:
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            try:
                age = now - lock_path.stat().st_mtime
            except FileNotFoundError:
                continue
            if age < LOCK_MAX_AGE_SECONDS:
                raise DescEditorBuildError("Збірка редактора вже виконується (lockfile присутній).")
            try:
                lock_path.unlink()
                LOGGER.warning("Виявлено застарілий lockfile для збірки редактора, видалено: %s", lock_path)
            except OSError as exc:
                raise DescEditorBuildError(
                    "Не вдалося звільнити застарілий lockfile збірки редактора."
                ) from exc
            continue
        else:
            with os.fdopen(fd, "w", encoding="utf-8") as fh:
                fh.write(json.dumps({"pid": os.getpid(), "timestamp": now}))
            return lock_path


def _release_build_lock(lock_path: Path) -> None:
    try:
        lock_path.unlink()
    except FileNotFoundError:
        return
    except Exception:
        LOGGER.exception("Не вдалося видалити lockfile збірки редактора: %s", lock_path)


def ensure_desc_editor_built(
    force: bool = False,
    *,
    install: bool = True,
    npm_executable: str | None = None,
    quiet: bool = False,
) -> Path:
    """Ensure that the description editor bundle is built and return the dist path."""
    lock_path: Path | None = None
    with _BUILD_LOCK:
        lock_path = _acquire_build_lock()
        try:
            return _ensure_desc_editor_built_impl(
                force=force,
                install=install,
                npm_executable=npm_executable,
                quiet=quiet,
            )
        finally:
            if lock_path:
                _release_build_lock(lock_path)


def _ensure_desc_editor_built_impl(
    force: bool = False,
    *,
    install: bool = True,
    npm_executable: str | None = None,
    quiet: bool = False,
) -> Path:
    """Internal implementation to allow lock handling around the build process."""
    if not EDITOR_ROOT.exists():
        raise DescEditorBuildError("Каталог desc-editor відсутній у репозиторії.")

    editor_sources = [
        EDITOR_ROOT / "package.json",
        EDITOR_ROOT / "package-lock.json",
        EDITOR_ROOT / "vite.config.ts",
        EDITOR_ROOT / "tsconfig.json",
        EDITOR_ROOT / "tsconfig.node.json",
        EDITOR_ROOT / "index.html",
        EDITOR_ROOT / "src",
    ]

    need_install = force or not INSTALL_STAMP.exists()
    if install and not need_install and editor_sources[1].exists():
        try:
            need_install = editor_sources[1].stat().st_mtime > INSTALL_STAMP.stat().st_mtime
        except FileNotFoundError:
            need_install = True

    need_build = force or not (DIST_DIR / "index.html").exists()
    if not need_build:
        try:
            dist_mtime = (DIST_DIR / "index.html").stat().st_mtime
        except FileNotFoundError:
            dist_mtime = 0.0
        src_mtime = _latest_mtime(editor_sources)
        if src_mtime > dist_mtime:
            need_build = True
        elif BUILD_STAMP.exists():
            try:
                need_build = editor_sources[1].stat().st_mtime > BUILD_STAMP.stat().st_mtime
            except FileNotFoundError:
                need_build = True

    npm_path: str | None = None

    check_runtime = (install and need_install) or need_build
    if check_runtime:
        if not quiet:
            print("[desc-editor] Перевіряємо Node.js…")
        _ensure_node_version()

    if install and need_install:
        npm_path = _resolve_npm(npm_executable)
        if not quiet:
            print("[desc-editor] Installing npm dependencies…")
        _run_command(
            [npm_path, "install", "--no-audit", "--no-fund"],
            cwd=EDITOR_ROOT,
            quiet=quiet,
        )
        INSTALL_STAMP.parent.mkdir(parents=True, exist_ok=True)
        INSTALL_STAMP.write_text(str(time.time()))
        need_build = True  # npm install may update deps

    if need_build:
        npm_path = npm_path or _resolve_npm(npm_executable)
        if not quiet:
            print("[desc-editor] Building production bundle…")
        _run_command([npm_path, "run", "build"], cwd=EDITOR_ROOT, quiet=quiet)
        BUILD_STAMP.parent.mkdir(parents=True, exist_ok=True)
        BUILD_STAMP.write_text(str(time.time()))

    if not (DIST_DIR / "index.html").exists():
        raise DescEditorBuildError(
            "Збірка редактора не створила index.html. Перевірте журнали npm run build."
        )

    return DIST_DIR


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Зібрати фронтенд редактора описів")
    parser.add_argument("--force", action="store_true", help="примусово перевстановити залежності та зібрати фронтенд")
    parser.add_argument("--skip-install", action="store_true", help="пропустити npm install")
    parser.add_argument("--npm", dest="npm_executable", help="шлях до npm, якщо не у PATH")
    parser.add_argument("--quiet", action="store_true", help="мінімізувати вивід команд")
    args = parser.parse_args(argv)

    try:
        ensure_desc_editor_built(
            force=args.force,
            install=not args.skip_install,
            npm_executable=args.npm_executable,
            quiet=args.quiet,
        )
    except DescEditorBuildError as exc:
        print(f"Помилка: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    raise SystemExit(main())
