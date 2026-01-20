"""Entry point for the Prom generator application."""
from __future__ import annotations

import logging
import sys
from pathlib import Path

from app_paths import get_data_dir, migrate_legacy_files, setup_logging
from errors import MissingDependencyError

LOGGER = logging.getLogger(__name__)


def _report_startup_error(message: str) -> None:
    """Show a user-friendly error during startup without crashing imports."""

    try:
        import tkinter as tk
        from tkinter import messagebox
    except Exception:
        print(message, file=sys.stderr)
        return

    root = None
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("ProdGen", message)
    except Exception:
        print(message, file=sys.stderr)
    finally:
        if root is not None:
            try:
                root.destroy()
            except Exception:
                pass


def main() -> int:
    data_dir = get_data_dir()
    data_dir.mkdir(parents=True, exist_ok=True)
    setup_logging()
    migrate_legacy_files(Path(__file__).resolve().parent)

    try:
        from database import init_db

        init_db()
        from ui.app import App
    except MissingDependencyError as exc:
        _report_startup_error(str(exc))
        return 1
    except Exception:  # pragma: no cover - defensive guard for unexpected failures
        LOGGER.exception("Помилка запуску застосунку")
        _report_startup_error("Не вдалося запустити застосунок. Перевірте журнали.")
        return 1

    try:
        app = App()
        app.mainloop()
    except MissingDependencyError as exc:
        _report_startup_error(str(exc))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
