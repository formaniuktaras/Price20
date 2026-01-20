"""CustomTkinter UI application for the Prom generator."""
from __future__ import annotations

import csv
import json
import logging
import os
import subprocess
import sys
import time
import queue
import re
import threading
import uuid
import webbrowser
from copy import deepcopy
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional, Sequence, Tuple, cast
from http import HTTPStatus

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from app_paths import get_data_dir, get_default_export_dir, get_logs_dir
from data_transfer import (
    DataTransferError,
    export_all_data_to_excel,
    import_all_data_from_excel,
)
from templates_service import (
    APP_TITLE,
    DEPENDENCY_WARNINGS,
    CATEGORY_SCOPE_DEFAULT_LABEL,
    GLOBAL_DESCRIPTION_KEY,
    DEFAULT_TEMPLATES,
    FILM_TYPE_DEFAULT_LABEL,
    TEMPLATE_LANGUAGE_DEFAULT_LABEL,
    ExportError,
    export_products,
    generate_export_rows,
    get_available_export_formats,
    load_export_fields,
    load_templates,
    load_title_tags_templates,
    resolve_title_tags,
    save_export_fields,
    save_templates,
    save_title_tags_templates,
    _title_tags_block,
    _normalize_language_definitions,
    _normalize_export_field_languages,
    _normalize_title_tags_block,
    _normalize_template_language_entry,
    _get_language_template_value,
    _set_language_template_value,
    _looks_like_formula,
    _copy_default_export_fields,
    _row_to_values,
    Template,
    TemplateError,
)

from database import (
    add_brand,
    add_category,
    add_model,
    delete_brand,
    delete_category,
    delete_model,
    delete_spec,
    get_brands,
    get_categories,
    get_models,
    get_specs,
    insert_spec,
    replace_specs,
    rename_brand,
    rename_category,
    rename_model,
    update_spec,
)

from formula_engine import FormulaEngine, FormulaError
from specs_io import format_specs_for_clipboard, parse_specs_payload

from desc_editor_build import DescEditorBuildError, ensure_desc_editor_built
from errors import MissingDependencyError
from settings_service import load_settings, save_settings, validate_settings
from ui.settings_dialog import SettingsDialog
from ui.theme_manager import ThemeManager

logger = logging.getLogger(__name__)

DESC_EDITOR_DIST = (Path(__file__).resolve().parent.parent / "desc-editor" / "dist").resolve()
DESC_EDITOR_ENTRY = DESC_EDITOR_DIST / "index.html"


try:
    import customtkinter as ctk
except ModuleNotFoundError as exc:
    raise MissingDependencyError(
        "Бібліотека CustomTkinter не знайдена. Встановіть її командою 'pip install customtkinter'."
    ) from exc

_INPUT_SPLIT_RE = re.compile(r"[\n\r,;\u201a\u201e\uFF0C\u3001]+")
def split_catalog_input(raw: str):
    if not raw:
        return []
    parts = [part.strip() for part in _INPUT_SPLIT_RE.split(raw) if part.strip()]
    unique = []
    seen = set()
    for part in parts:
        if part not in seen:
            unique.append(part)
            seen.add(part)
    return unique

  
def create_inline_entry(parent, text: str, theme_colors: Optional[Dict[str, str]] = None):
    entry = tk.Entry(parent)
    font = ctk.CTkFont()
    entry.configure(font=font)
    entry._ctk_font = font  # keep reference to avoid garbage collection
    mode = (ctk.get_appearance_mode() or "light").lower()
    if theme_colors:
        bg = theme_colors.get("widget_fg", "#2b2b2b")
        fg = theme_colors.get("text", "#f2f2f2")
        border = theme_colors.get("border", "#565b5e")
        selection_bg = theme_colors.get("selection_bg", "#1f6aa5")
        selection_fg = theme_colors.get("selection_text", fg)
        caret = theme_colors.get("caret", fg)
    elif mode == "dark":
        bg = "#2b2b2b"
        fg = "#f2f2f2"
        border = "#565b5e"
        selection_bg = "#1f6aa5"
        selection_fg = "#ffffff"
        caret = fg
    else:
        bg = "#ffffff"
        fg = "#1f1f1f"
        border = "#a5a5a5"
        selection_bg = "#1f6aa5"
        selection_fg = "#ffffff"
        caret = fg
    entry.configure(
        background=bg,
        foreground=fg,
        insertbackground=caret,
        selectbackground=selection_bg,
        selectforeground=selection_fg,
        highlightthickness=1,
        highlightbackground=border,
        highlightcolor=selection_bg,
        borderwidth=0,
        relief="flat",
    )
    entry.insert(0, text)
    entry.select_range(0, tk.END)
    entry.focus_set()
    return entry

def show_error(msg: str):
    messagebox.showerror("Помилка", msg)

def show_info(msg: str):
    messagebox.showinfo("Інформація", msg)


class DescriptionEditorError(RuntimeError):
    """Raised when the browser-based description editor cannot be launched."""


class DescriptionEditorHost:
    """Bridge between the desktop UI and the web-based description editor."""

    def __init__(self, docs: Dict[str, Dict[str, object]], active_lang: str = "uk") -> None:
        self._docs = docs
        self._active_lang = active_lang if active_lang in docs else next(iter(docs or {"uk": {}}))
        self._server: Optional[ThreadingHTTPServer] = None
        self._thread: Optional[threading.Thread] = None
        self._session_id = uuid.uuid4().hex
        self._result_event = threading.Event()
        self.result: Optional[Dict[str, Dict[str, object]]] = None

    # ------------------------- HTTP server helpers -------------------------
    def _shutdown_server(self) -> None:
        if self._server is not None:
            try:
                self._server.shutdown()
            except Exception:  # pragma: no cover - best effort cleanup
                logger.exception("Не вдалося коректно зупинити веб-сервер редактора")
            try:
                self._server.server_close()
            except Exception:
                logger.exception("Не вдалося закрити сокет веб-сервера")
            self._server = None
        if self._thread is not None:
            self._thread.join(timeout=2)
            self._thread = None

    def _api_payload(self) -> Dict[str, object]:
        return {
            "activeLang": self._active_lang,
            "docs": self._docs,
        }

    def _accept_result(self, payload: Dict[str, object]) -> bool:
        docs = payload.get("docs") if isinstance(payload, dict) else None
        if not isinstance(docs, dict):
            return False
        active_lang = payload.get("activeLang")
        if isinstance(active_lang, str) and active_lang:
            self._active_lang = active_lang
        self.result = docs  # type: ignore[assignment]
        self._result_event.set()
        return True

    def _create_handler(self):
        host = self

        class _RequestHandler(SimpleHTTPRequestHandler):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, directory=str(DESC_EDITOR_DIST), **kwargs)

            def log_message(self, format: str, *args: object) -> None:  # pragma: no cover - noisy
                logger.debug("desc-editor: " + format, *args)

            def _send_json(self, data: Dict[str, object], status: int = HTTPStatus.OK) -> None:
                body = json.dumps(data).encode("utf-8")
                self.send_response(status)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def _send_html(self, html: str, status: int = HTTPStatus.OK) -> None:
                body = html.encode("utf-8")
                self.send_response(status)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def do_GET(self) -> None:  # noqa: N802 - required name
                if self.path.startswith(f"/api/session/{host._session_id}/state"):
                    self._send_json(host._api_payload())
                    return
                if self.path.startswith("/close"):
                    self._send_html(
                        """
<!DOCTYPE html>
<html lang="uk">
<head><meta charset="utf-8"><title>Редактор опису</title></head>
<body style="font-family: sans-serif; margin: 40px;">
  <h1>Редактор опису</h1>
  <p>Дані передано у застосунок. Ви можете закрити цю вкладку.</p>
</body>
</html>
""".strip()
                    )
                    return
                super().do_GET()

            def do_POST(self) -> None:  # noqa: N802 - required name
                if self.path.startswith(f"/api/session/{host._session_id}/save"):
                    length = int(self.headers.get("Content-Length") or 0)
                    raw = self.rfile.read(length)
                    try:
                        payload = json.loads(raw.decode("utf-8"))
                    except json.JSONDecodeError:
                        self._send_json({"error": "Invalid JSON"}, status=HTTPStatus.BAD_REQUEST)
                        return
                    if not host._accept_result(payload):
                        self._send_json({"error": "Invalid payload"}, status=HTTPStatus.BAD_REQUEST)
                        return
                    self._send_json({"status": "ok"})
                    return
                self.send_error(HTTPStatus.NOT_FOUND, "Endpoint not found")

        return _RequestHandler

    # ------------------------- Public API -------------------------
    def launch(self) -> None:
        try:
            ensure_desc_editor_built()
        except DescEditorBuildError as exc:
            raise DescriptionEditorError(
                "Не вдалося підготувати фронтенд редактора опису. "
                "Переконайтеся, що встановлено Node.js з npm та повторіть спробу."
            ) from exc

        if not DESC_EDITOR_ENTRY.exists():
            raise DescriptionEditorError(
                "Фронтенд редактора не знайдено навіть після автоматичної збірки. "
                "Перевірте журнал виконання npm run build."
            )

        handler_cls = self._create_handler()
        server = ThreadingHTTPServer(("127.0.0.1", 0), handler_cls)
        server.daemon_threads = True
        self._server = server
        self._thread = threading.Thread(target=server.serve_forever, daemon=True)
        self._thread.start()

        port = server.server_address[1]
        url = f"http://127.0.0.1:{port}/index.html?session={self._session_id}"
        opened = webbrowser.open(url, new=1)
        if not opened:
            logger.warning("Не вдалося автоматично відкрити браузер для редактора опису")

    def poll_result(self) -> Optional[Dict[str, Dict[str, object]]]:
        if self._result_event.is_set():
            return cast(Optional[Dict[str, Dict[str, object]]], self.result)
        return None

    def wait(self, timeout: Optional[float] = None) -> Optional[Dict[str, Dict[str, object]]]:
        if self._result_event.wait(timeout):
            return cast(Optional[Dict[str, Dict[str, object]]], self.result)
        return None

    @property
    def is_running(self) -> bool:
        return self._server is not None

    def close(self) -> None:
        self._shutdown_server()


class SpecsBulkEditor(ctk.CTkToplevel):
    """Simple text-based editor for bulk specification updates."""

    def __init__(
        self,
        master,
        specs: Sequence[Tuple[str, str]],
        apply_callback,
    ) -> None:
        super().__init__(master)
        self._apply_callback = apply_callback
        self.title("Масове редагування характеристик")
        self.geometry("640x520")
        self.minsize(520, 360)
        self.transient(master)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        info_text = (
            "Кожен рядок відповідає одній характеристиці. "
            "Використовуйте табуляцію або ; для розділення назви та значення."
        )
        info = ctk.CTkLabel(self, text=info_text, wraplength=580, justify="left")
        info.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 8))

        self.textbox = ctk.CTkTextbox(self, wrap="none")
        self.textbox.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 12))
        normalized_specs: List[Tuple[str, str]] = []
        for key, value in specs:
            safe_key = "" if key is None else str(key)
            safe_value = "" if value is None else str(value)
            normalized_specs.append((safe_key, safe_value))
        initial = format_specs_for_clipboard(normalized_specs)
        if initial:
            self.textbox.insert("1.0", initial)
        self.textbox.focus_set()

        btn_frame = ctk.CTkFrame(self)
        btn_frame.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 16))
        btn_frame.columnconfigure(0, weight=1)

        self.apply_button = ctk.CTkButton(btn_frame, text="Застосувати", command=self._on_apply)
        self.apply_button.grid(row=0, column=1, padx=(0, 8))

        self.cancel_button = ctk.CTkButton(
            btn_frame,
            text="Скасувати",
            fg_color="#444444",
            hover_color="#333333",
            command=self._on_cancel,
        )
        self.cancel_button.grid(row=0, column=2)

    def _on_apply(self) -> None:
        if not callable(self._apply_callback):
            self.destroy()
            return
        payload = self.textbox.get("1.0", tk.END)
        should_close = self._apply_callback(payload)
        if should_close:
            self.destroy()

    def _on_cancel(self) -> None:
        self.destroy()


# ============================ GUI: ВІКНО ХАРАКТЕРИСТИК ============================

class SpecsWindow(ctk.CTkToplevel):
    def __init__(self, master, model_id: int, model_name: str):
        super().__init__(master)
        self.model_id = model_id
        self.title(f"Характеристики: {model_name}")
        self.geometry("700x480")
        self.resizable(True, True)

        self._tree_style_name, self._tree_colors = self._init_tree_style()

        self._current_specs: List[Tuple[int, str, Optional[str]]] = []
        self._filter_after_id: Optional[str] = None
        self._bulk_editor = None

        binder = getattr(master, "_bind_clipboard_shortcuts", None)

        # Панель пошуку
        filter_bar = ctk.CTkFrame(self)
        filter_bar.pack(fill="x", padx=10, pady=(10, 0))

        ctk.CTkLabel(filter_bar, text="Пошук:").pack(side="left", padx=(6, 6))

        self.filter_var = tk.StringVar()
        self.filter_entry = ctk.CTkEntry(
            filter_bar,
            textvariable=self.filter_var,
            placeholder_text="Наприклад: Діагональ",
        )
        self.filter_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.filter_var.trace_add("write", lambda *_: self._schedule_filter_update())
        self.filter_entry.bind("<Return>", lambda _e: self._apply_filter())
        self.filter_entry.bind("<Escape>", lambda _e: self._clear_filter())
        if callable(binder):
            binder(self.filter_entry)

        self.clear_filter_button = ctk.CTkButton(
            filter_bar,
            text="Скинути",
            width=90,
            command=self._clear_filter,
        )
        self.clear_filter_button.pack(side="left")

        # Таблиця
        self.tree = ttk.Treeview(
            self,
            columns=("key", "value"),
            show="headings",
            height=16,
            style=self._tree_style_name,
        )
        self.tree.heading("key", text="Назва параметра")
        self.tree.heading("value", text="Значення")
        self.tree.column("key", width=260, anchor="w")
        self.tree.column("value", width=360, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=10, pady=(6, 10))

        scroll = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        scroll.place(relx=1.0, rely=0.0, relheight=1.0, anchor="ne")
        self.tree.bind("<Delete>", self._on_delete_key)
        self.tree.bind("<Button-1>", self._on_tree_click, add="+")
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        self._rename_click = (None, None, 0.0)
        self._rename_entry = None
        self._rename_meta = None
        self._rename_delay_min = 0.35
        self._rename_delay_max = 4.0

        # Панель керування
        ctrl = ctk.CTkFrame(self)
        ctrl.pack(fill="x", padx=10, pady=(0,10))

        self.key_entry = ctk.CTkEntry(ctrl, placeholder_text="Напр.: Діагональ екрану")
        self.key_entry.pack(side="left", fill="x", expand=True, padx=5)
        if callable(binder):
            binder(self.key_entry)
        self.val_entry = ctk.CTkEntry(ctrl, placeholder_text="Напр.: 6.7''")
        self.val_entry.pack(side="left", fill="x", expand=True, padx=5)
        if callable(binder):
            binder(self.val_entry)

        self.add_button = ctk.CTkButton(ctrl, text="Додати", command=self._add)
        self.add_button.pack(side="left", padx=5)
        self.update_button = ctk.CTkButton(ctrl, text="Оновити", command=self._edit)
        self.update_button.pack(side="left", padx=5)
        self.delete_button = ctk.CTkButton(
            ctrl,
            text="Видалити",
            fg_color="#8b0000",
            hover_color="#a40000",
            command=self._delete,
        )
        self.delete_button.pack(side="left", padx=5)

        bulk_ctrl = ctk.CTkFrame(self)
        bulk_ctrl.pack(fill="x", padx=10, pady=(0, 10))

        self.bulk_edit_button = ctk.CTkButton(
            bulk_ctrl,
            text="Масове редагування",
            command=self._open_bulk_editor,
        )
        self.bulk_edit_button.pack(side="left", padx=5)

        self.import_clipboard_button = ctk.CTkButton(
            bulk_ctrl,
            text="Імпорт з буфера",
            command=self._import_from_clipboard,
        )
        self.import_clipboard_button.pack(side="left", padx=5)

        self.import_file_button = ctk.CTkButton(
            bulk_ctrl,
            text="Імпорт з файлу",
            command=self._import_from_file,
        )
        self.import_file_button.pack(side="left", padx=5)

        self.export_clipboard_button = ctk.CTkButton(
            bulk_ctrl,
            text="Експорт у буфер",
            command=self._export_to_clipboard,
        )
        self.export_clipboard_button.pack(side="left", padx=5)

        self.export_file_button = ctk.CTkButton(
            bulk_ctrl,
            text="Експорт у файл",
            command=self._export_to_file,
        )
        self.export_file_button.pack(side="left", padx=5)

        self._refresh()

    def _on_delete_key(self, _event):
        self._delete()
        return "break"

    def _init_tree_style(self):
        style = ttk.Style(self)
        style_name = "Specs.Treeview"
        theme_colors = {}
        master = getattr(self, "master", None)
        if master is not None:
            theme_manager = getattr(master, "theme_manager", None)
            if theme_manager is not None:
                theme_colors = getattr(theme_manager, "colors", {}) or {}
        mode = (ctk.get_appearance_mode() or "light").lower()
        if theme_colors:
            bg = theme_colors.get("widget_fg", "#1f1f1f")
            alt_bg = theme_colors.get("surface", bg)
            fg = theme_colors.get("text", "#f2f2f2")
            border = theme_colors.get("border", "#565b5e")
            heading_bg = theme_colors.get("header_bg", bg)
            heading_fg = theme_colors.get("header_text", fg)
            heading_border = theme_colors.get("header_border", border)
            hover_bg = theme_colors.get("header_border", border)
            select_bg = theme_colors.get("selection_bg", "#1f6aa5")
            select_fg = theme_colors.get("selection_text", "#ffffff")
        elif mode == "dark":
            bg = "#1f1f1f"
            alt_bg = "#242424"
            fg = "#f2f2f2"
            border = "#565b5e"
            heading_bg = "#232323"
            heading_fg = fg
            heading_border = border
            hover_bg = "#303030"
            select_bg = "#1f6aa5"
            select_fg = "#ffffff"
        else:
            bg = "#ffffff"
            alt_bg = "#f5f5f5"
            fg = "#1f1f1f"
            border = "#a5a5a5"
            heading_bg = "#f1f1f1"
            heading_fg = fg
            heading_border = border
            hover_bg = "#e2e2e2"
            select_bg = "#1f6aa5"
            select_fg = "#ffffff"
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        def safe_style_configure(style_name: str, **kwargs: object) -> None:
            try:
                style.configure(style_name, **kwargs)
            except Exception:
                logger.exception("Не вдалося налаштувати ttk стиль %s", style_name)

        def safe_style_map(style_name: str, **kwargs: object) -> None:
            try:
                style.map(style_name, **kwargs)
            except Exception:
                logger.exception("Не вдалося налаштувати ttk map для %s", style_name)

        safe_style_configure(
            style_name,
            background=bg,
            foreground=fg,
            fieldbackground=bg,
            bordercolor=border,
            borderwidth=1,
            rowheight=28,
            relief="flat",
        )
        safe_style_map(
            style_name,
            background=[("selected", select_bg)],
            foreground=[("selected", select_fg)],
        )
        heading_style = f"{style_name}.Heading"
        safe_style_configure(
            heading_style,
            background=heading_bg,
            foreground=heading_fg,
            bordercolor=heading_border,
            relief="flat",
        )
        safe_style_map(heading_style, background=[("active", hover_bg)])
        return style_name, {
            "row_even": bg,
            "row_odd": alt_bg,
            "fg": fg,
            "select_fg": select_fg,
        }

    def _refresh(self):
        if self._rename_entry is not None:
            self._finish_inline_edit(save=False)
        fresh_specs: List[Tuple[int, str, Optional[str]]] = []
        for sid, key, value in get_specs(self.model_id):
            if not isinstance(key, str):
                key = str(key)
            if value is not None and not isinstance(value, str):
                value = str(value)
            fresh_specs.append((sid, key, value))
        self._current_specs = fresh_specs
        self._apply_filter()

    def _on_tree_click(self, event):
        row = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        region = self.tree.identify_region(event.x, event.y)
        now = time.time()
        if region not in {"cell", "tree"}:
            self._rename_click = (None, None, now)
            self.after_idle(self._update_controls_from_selection)
            return
        if not row or column not in {"#1", "#2"}:
            self._rename_click = (row, column, now)
            self.after_idle(self._update_controls_from_selection)
            return
        last_row, last_col, last_time = self._rename_click
        self._rename_click = (row, column, now)
        delay = now - last_time
        if row == last_row and column == last_col and self._rename_delay_min <= delay <= self._rename_delay_max:
            self.after(0, lambda: self._start_inline_edit(row, column))
        self.after_idle(self._update_controls_from_selection)

    def _start_inline_edit(self, iid: str, column: str):
        if not self.tree.exists(iid):
            return
        if column not in {"#1", "#2"}:
            return
        values = self.tree.item(iid, "values")
        if not values:
            return
        index = 0 if column == "#1" else 1
        original = values[index]
        bbox = self.tree.bbox(iid, column)
        if not bbox:
            return
        if self._rename_entry is not None:
            self._finish_inline_edit(save=False)
        theme_colors = None
        master = getattr(self, "master", None)
        if master is not None:
            theme_manager = getattr(master, "theme_manager", None)
            if theme_manager is not None:
                theme_colors = getattr(theme_manager, "colors", None)
        entry = create_inline_entry(self.tree, original, theme_colors=theme_colors)
        x, y, width, height = bbox
        entry.place(x=x, y=y, width=width, height=height)
        field = "key" if column == "#1" else "value"
        self._rename_entry = entry
        self._rename_meta = (iid, field)
        entry.bind("<Return>", lambda _e: self._finish_inline_edit(save=True))
        entry.bind("<KP_Enter>", lambda _e: self._finish_inline_edit(save=True))
        entry.bind("<Escape>", lambda _e: self._finish_inline_edit(save=False))
        entry.bind("<FocusOut>", lambda _e: self._finish_inline_edit(save=True))

    def _finish_inline_edit(self, save: bool):
        if not self._rename_entry or not self._rename_meta:
            return
        entry = self._rename_entry
        iid, field = self._rename_meta
        self._rename_entry = None
        self._rename_meta = None
        new_value = entry.get().strip()
        entry.destroy()
        if not save:
            self._restore_selection(iid)
            return
        if field == "key" and not new_value:
            show_error("Назва параметра не може бути порожньою.")
            self._restore_selection(iid)
            return
        sid = int(iid.split("_")[1])
        values = self.tree.item(iid, "values")
        if not values or len(values) < 2:
            self._restore_selection(iid)
            return
        key, value = values[0], values[1]
        if field == "key":
            if new_value == key:
                self._restore_selection(iid)
                return
            key = new_value
        else:
            if new_value == value:
                self._restore_selection(iid)
                return
            value = new_value
        try:
            update_spec(sid, key, value)
        except ValueError as exc:
            show_error(str(exc))
            self._refresh()
            self.after(10, lambda: self._restore_selection(f"spec_{sid}"))
            return
        self._refresh()
        self.after(10, lambda: self._restore_selection(f"spec_{sid}"))

    def _restore_selection(self, iid: str):
        def _select():
            if not self.tree.exists(iid):
                return
            self.tree.selection_set(iid)
            self.tree.focus(iid)
            self.tree.see(iid)
        self.after(10, _select)
        self.after(15, self._update_controls_from_selection)

    def _add(self):
        k = self.key_entry.get().strip()
        v = self.val_entry.get().strip()
        if not k:
            show_error("Введіть назву параметра.")
            return
        insert_spec(self.model_id, k, v)
        self.key_entry.delete(0, tk.END); self.val_entry.delete(0, tk.END)
        self._refresh()

    def _edit(self):
        sel = self.tree.selection()
        if not sel:
            show_error("Оберіть рядок у таблиці.")
            return
        sid = int(sel[0].split("_")[1])
        k = self.key_entry.get().strip()
        v = self.val_entry.get().strip()
        if not k:
            show_error("Введіть назву параметра.")
            return
        try:
            update_spec(sid, k, v)
        except ValueError as exc:
            show_error(str(exc))
            self._refresh()
            self._restore_selection(f"spec_{sid}")
            return
        self._refresh()
        self._restore_selection(f"spec_{sid}")

    def _delete(self):
        selection = list(self.tree.selection())
        if not selection:
            show_error("Оберіть рядок у таблиці.")
            return
        prompt = (
            "Видалити вибрану характеристику?"
            if len(selection) == 1
            else "Видалити вибрані характеристики?"
        )
        if not messagebox.askyesno("Підтвердження", prompt):
            return
        ids = [int(iid.split("_")[1]) for iid in selection]
        for sid in ids:
            delete_spec(sid)
        self._refresh()
        self._update_controls_from_selection()

    def _on_tree_select(self, _event):
        self._update_controls_from_selection()

    def _update_controls_from_selection(self):
        selection = list(self.tree.selection())
        if len(selection) != 1:
            self.key_entry.delete(0, tk.END)
            self.val_entry.delete(0, tk.END)
            self._set_button_state(self.update_button, False)
        else:
            values = self.tree.item(selection[0], "values")
            if values:
                self.key_entry.delete(0, tk.END)
                self.key_entry.insert(0, values[0])
                self.val_entry.delete(0, tk.END)
                self.val_entry.insert(0, values[1] if len(values) > 1 else "")
            self._set_button_state(self.update_button, True)
        self._set_button_state(self.delete_button, bool(selection))

    @staticmethod
    def _set_button_state(button, enabled: bool):
        if button is None:
            return
        button.configure(state="normal" if enabled else "disabled")

    def _schedule_filter_update(self):
        if self._filter_after_id is not None:
            try:
                self.after_cancel(self._filter_after_id)
            except Exception:
                pass
        self._filter_after_id = self.after(150, self._apply_filter)

    def _apply_filter(self):
        if self._filter_after_id is not None:
            try:
                self.after_cancel(self._filter_after_id)
            except Exception:
                pass
        self._filter_after_id = None
        self._render_specs(self._filtered_specs())

    def _clear_filter(self):
        if self.filter_var.get():
            self.filter_var.set("")
        self._apply_filter()
        try:
            self.filter_entry.focus_set()
        except Exception:
            pass

    def _filtered_specs(self) -> List[Tuple[int, str, Optional[str]]]:
        query = (self.filter_var.get() or "").strip().lower()
        if not query:
            return list(self._current_specs)
        filtered: List[Tuple[int, str, Optional[str]]] = []
        for sid, key, value in self._current_specs:
            key_text = key or ""
            value_text = ""
            if value is not None:
                value_text = str(value)
            if query in key_text.lower() or (value_text and query in value_text.lower()):
                filtered.append((sid, key_text, value))
        return filtered

    def _render_specs(self, specs: Optional[Sequence[Tuple[int, str, Optional[str]]]] = None):
        data = list(specs or [])
        if not data:
            data = [] if specs is not None else list(self._current_specs)
        selected = list(self.tree.selection())
        focus = self.tree.focus()
        self.tree.delete(*self.tree.get_children())
        for idx, (sid, key, value) in enumerate(data):
            iid = f"spec_{sid}"
            display_value = value if value is not None else ""
            tags = ("even",) if idx % 2 == 0 else ("odd",)
            self.tree.insert("", "end", iid=iid, values=(key, display_value), tags=tags)
        if self._tree_colors:
            self.tree.tag_configure(
                "even",
                background=self._tree_colors.get("row_even"),
                foreground=self._tree_colors.get("fg"),
            )
            self.tree.tag_configure(
                "odd",
                background=self._tree_colors.get("row_odd"),
                foreground=self._tree_colors.get("fg"),
            )
        restored = [iid for iid in selected if self.tree.exists(iid)]
        if restored:
            self.tree.selection_set(restored)
            if focus and self.tree.exists(focus):
                self.tree.focus(focus)
        else:
            self.tree.selection_remove(self.tree.selection())
            self.tree.focus("")
        self.after_idle(self._update_controls_from_selection)

    def _collect_specs(self) -> List[Tuple[str, str]]:
        specs: List[Tuple[str, str]] = []
        for _sid, key, value in self._current_specs:
            safe_key = (key or "").strip()
            safe_value = ""
            if value is not None:
                safe_value = str(value).strip()
            specs.append((safe_key, safe_value))
        return specs

    def _open_bulk_editor(self):
        if self._bulk_editor is not None and self._bulk_editor.winfo_exists():
            try:
                self._bulk_editor.focus_set()
            except Exception:
                pass
            return
        specs_payload: List[Tuple[str, str]] = []
        for _sid, key, value in self._current_specs:
            safe_key = str(key) if key is not None else ""
            safe_value = "" if value is None else str(value)
            specs_payload.append((safe_key, safe_value))
        editor = SpecsBulkEditor(self, specs_payload, self._apply_bulk_editor_payload)
        binder = getattr(self.master, "_bind_clipboard_shortcuts", None)
        if callable(binder):
            binder(editor.textbox)
        editor.bind("<Destroy>", lambda _e: setattr(self, "_bulk_editor", None))
        self._bulk_editor = editor

    def _apply_bulk_editor_payload(self, raw: str) -> bool:
        pairs = parse_specs_payload(raw)
        ordered: List[Tuple[str, str]] = []
        index_map: Dict[str, int] = {}
        for key, value in pairs:
            normalized_key = key.strip()
            if not normalized_key:
                continue
            normalized_value = (value or "").strip()
            lookup = normalized_key.lower()
            existing_index = index_map.get(lookup)
            if existing_index is None:
                index_map[lookup] = len(ordered)
                ordered.append((normalized_key, normalized_value))
            else:
                ordered[existing_index] = (normalized_key, normalized_value)

        existing_map: Dict[str, Tuple[Optional[int], str, str]] = {}
        for sid, key, value in self._current_specs:
            normalized_key = (key or "").strip()
            if not normalized_key:
                continue
            lookup = normalized_key.lower()
            existing_map[lookup] = (sid, (value or "").strip(), normalized_key)

        if not ordered and existing_map:
            if not messagebox.askyesno(
                "Підтвердження",
                "Очистити всі характеристики моделі?",
            ):
                return False

        after_values = {key.lower(): value for key, value in ordered}
        after_keys = {key.lower(): key for key, _value in ordered}

        inserted = sum(1 for key in after_values if key not in existing_map)
        updated = 0
        for key, value in after_values.items():
            if key not in existing_map:
                continue
            _, current_value, current_key = existing_map[key]
            new_key = after_keys.get(key, current_key)
            if current_value != value or current_key != new_key:
                updated += 1
        removed = sum(1 for key in existing_map if key not in after_values)

        try:
            replace_specs(self.model_id, ordered)
        except Exception as exc:
            logger.exception("Failed to apply bulk specs", exc_info=exc)
            show_error(f"Не вдалося зберегти зміни: {exc}")
            return False

        self._refresh()

        summary_parts = []
        if inserted:
            summary_parts.append(f"додано: {inserted}")
        if updated:
            summary_parts.append(f"оновлено: {updated}")
        if removed:
            summary_parts.append(f"видалено: {removed}")
        if not summary_parts:
            summary_parts.append("змін не внесено")
        show_info("Масове редагування виконано:\n" + ", ".join(summary_parts))
        return True

    # ------------------------------ Імпорт/експорт ---------------------------------

    def _import_from_clipboard(self):
        try:
            raw = self.clipboard_get()
        except tk.TclError:
            show_error("Не вдалося прочитати буфер обміну.")
            return
        self._apply_import_payload(raw, source="буфера обміну")

    def _import_from_file(self):
        path = filedialog.askopenfilename(
            title="Імпорт характеристик",
            filetypes=(
                ("Текстові файли", "*.txt"),
                ("CSV/TSV", "*.csv *.tsv"),
                ("Усі файли", "*.*"),
            ),
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8-sig") as fh:
                raw = fh.read()
        except OSError as exc:
            show_error(f"Не вдалося відкрити файл:\n{exc}")
            return
        self._apply_import_payload(raw, source=os.path.basename(path))

    def _apply_import_payload(self, raw: str, source: str) -> None:
        pairs = parse_specs_payload(raw)
        if not pairs:
            show_info("Не знайдено характеристик для імпорту.")
            return
        existing: Dict[str, Tuple[Optional[int], str, str]] = {}
        for sid, key, value in get_specs(self.model_id):
            normalized_key = (key or "").strip()
            if not normalized_key:
                continue
            existing[normalized_key.lower()] = (sid, (value or "").strip(), normalized_key)
        inserted = 0
        updated = 0
        skipped = 0
        for key, value in pairs:
            normalized_key = key.strip()
            if not normalized_key:
                skipped += 1
                continue
            normalized_value = (value or "").strip()
            lookup = normalized_key.lower()
            stored = existing.get(lookup)
            if stored is None:
                new_id = insert_spec(self.model_id, normalized_key, normalized_value)
                existing[lookup] = (new_id, normalized_value, normalized_key)
                inserted += 1
                continue
            sid, current_value, current_key = stored
            if current_value == normalized_value and current_key == normalized_key:
                skipped += 1
                continue
            if sid is not None:
                try:
                    update_spec(sid, normalized_key, normalized_value)
                except ValueError as exc:
                    show_error(str(exc))
                    continue
                else:
                    existing[lookup] = (sid, normalized_value, normalized_key)
                    updated += 1
            else:
                new_id = insert_spec(self.model_id, normalized_key, normalized_value)
                existing[lookup] = (new_id, normalized_value, normalized_key)
                inserted += 1
        self._refresh()
        summary_parts = []
        if inserted:
            summary_parts.append(f"додано: {inserted}")
        if updated:
            summary_parts.append(f"оновлено: {updated}")
        if skipped:
            summary_parts.append(f"без змін: {skipped}")
        if not summary_parts:
            summary_parts.append("змін не внесено")
        show_info(f"Імпорт завершено ({source}):\n" + ", ".join(summary_parts))

    def _export_to_clipboard(self):
        specs = self._collect_specs()
        if not specs:
            show_info("Немає характеристик для експорту.")
            return
        payload = format_specs_for_clipboard(specs)
        try:
            self.clipboard_clear()
            self.clipboard_append(payload)
        except tk.TclError:
            show_error("Не вдалося записати дані у буфер обміну.")
            return
        show_info("Характеристики скопійовано до буфера обміну.")

    def _export_to_file(self):
        specs = get_specs(self.model_id)
        if not specs:
            show_info("Немає характеристик для експорту.")
            return
        path = filedialog.asksaveasfilename(
            title="Експорт характеристик",
            defaultextension=".csv",
            filetypes=(
                ("CSV файл", "*.csv"),
                ("TSV файл", "*.tsv"),
                ("Текстовий файл", "*.txt"),
                ("Усі файли", "*.*"),
            ),
        )
        if not path:
            return
        delimiter = ";" if path.lower().endswith(".csv") else "\t"
        try:
            with open(path, "w", newline="", encoding="utf-8") as fh:
                writer = csv.writer(fh, delimiter=delimiter)
                writer.writerow(["Назва параметра", "Значення"])
                for _sid, key, value in specs:
                    writer.writerow([key, value or ""])
        except OSError as exc:
            show_error(f"Не вдалося зберегти файл:\n{exc}")
            return
        show_info("Характеристики збережено у файл.")

# ============================ GUI: ОСНОВНИЙ ДОДАТОК ============================

class _TkAppCompatProxy:
    """Proxy that allows attaching arbitrary Python attributes to a tkapp."""

    __slots__ = ("_tkapp", "_extras")

    def __init__(self, tkapp):
        object.__setattr__(self, "_tkapp", tkapp)
        object.__setattr__(self, "_extras", {})

    def __getattr__(self, name):
        extras = object.__getattribute__(self, "_extras")
        if name in extras:
            return extras[name]
        return getattr(object.__getattribute__(self, "_tkapp"), name)

    def __setattr__(self, name, value):
        if name in {"_tkapp", "_extras"}:
            object.__setattr__(self, name, value)
            return
        extras = object.__getattribute__(self, "_extras")
        extras[name] = value

    def __delattr__(self, name):
        if name in {"_tkapp", "_extras"}:
            raise AttributeError(name)
        extras = object.__getattribute__(self, "_extras")
        if name in extras:
            del extras[name]
            return
        delattr(object.__getattribute__(self, "_tkapp"), name)

    def __dir__(self):
        extras = object.__getattribute__(self, "_extras")
        base_dir = dir(object.__getattribute__(self, "_tkapp"))
        return sorted(set(base_dir).union(extras.keys()))


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        raw_tkapp = object.__getattribute__(self, "tk")
        if not isinstance(raw_tkapp, _TkAppCompatProxy):
            object.__setattr__(self, "tk", _TkAppCompatProxy(raw_tkapp))
        tkapp = object.__getattribute__(self, "tk")
        self.title(APP_TITLE)
        self.geometry("1100x680")
        self.minsize(980, 640)
        self.settings = load_settings()
        self.theme_manager = ThemeManager(self)
        self.theme_manager.register(self, "background")
        self.theme_manager.apply(self.settings, apply_widgets=False)
        ctk.set_default_color_theme("blue")

        self._setup_ttk_styles()

        self._generation_task_running = False
        self._active_generation_thread: Optional[threading.Thread] = None
        self._progress_lock = threading.Lock()
        self._last_progress_update = 0.0
        self._ui_event_queue: "queue.Queue[Callable[[], None]]" = queue.Queue()
        self._ui_queue_job: Optional[str] = None

        self.templates = load_templates()
        self.title_tags_templates = load_title_tags_templates(self.templates)
        self.export_fields = load_export_fields()
        self.current_category_id = None
        self.current_brand_id = None
        self._current_film_type_key = None
        self._current_desc_category = None
        self._current_template_category = None
        self._template_category_label_to_key = {}
        self._template_category_key_to_label = {}
        self._template_film_label_to_key = {}
        self._template_film_key_to_label = {}
        self._template_language_label_to_code = {}
        self._template_language_code_to_label = {}
        self._current_template_language = None
        self._current_language_index = None
        self._gen_tree = None
        self._gen_tree_states = {}
        self._gen_tree_meta = {}
        self._gen_tree_labels = {}
        self.gen_filter_header = None
        self.gen_filter_panel = None
        self.gen_filter_toggle = None
        self.gen_filter_clear = None
        self.gen_filter_start_date_var = None
        self.gen_filter_start_time_var = None
        self.gen_filter_end_date_var = None
        self.gen_filter_end_time_var = None
        self.gen_filter_start_date_entry = None
        self.gen_filter_start_time_entry = None
        self.gen_filter_end_date_entry = None
        self.gen_filter_end_time_entry = None
        self.gen_filter_apply = None
        self._gen_filter_range: Tuple[Optional[datetime], Optional[datetime]] = (None, None)
        self._gen_filter_visible = False
        self._rename_clicks = {
            "cat": (None, 0.0),
            "brand": (None, 0.0),
            "model": (None, 0.0),
        }
        self._rename_entry = None
        self._rename_entry_meta = None
        self._rename_delay_min = 0.35
        self._rename_delay_max = 4.0
        self._export_selected_index = None
        self._export_tree_updating = False
        self._export_unknown_language_codes = []
        self.export_language_vars = []
        self.progress_bar = None
        self.progress_label = None
        self._preview_window = None
        self._active_desc_host = None
        self._desc_editor_prepare_thread = None
        self._desc_editor_ready = DESC_EDITOR_ENTRY.exists()
        self._desc_editor_retry_visible = False
        self._desc_editor_error_shown = False
        self.generate_preview_button = None
        self.generate_run_button = None
        self.generate_tree_buttons: List[ctk.CTkButton] = []
        self.choose_folder_button = None
        self.out_folder_entry = None
        self.file_menu_button = None
        self._file_menu = None
        self._film_type_vars: Dict[str, tk.BooleanVar] = {}
        self._film_type_cbs: Dict[str, ctk.CTkCheckBox] = {}
        self._film_type_order: List[str] = []
        self._film_types_cols = None
        self._film_types_scroll = None
        self._film_layout_job = None
        # Compatibility: some flows expect the filmtype name variable to exist during tab
        # construction even if the dedicated film type tab is hidden. Older widgets access
        # the variable through the low-level Tk interpreter (self.tk), so expose it there
        # via a proxy that can store Python attributes.
        self.filmtype_name_var = tk.StringVar(master=self, value="")
        tkapp.filmtype_name_var = self.filmtype_name_var

        self.filmtype_enabled_var = tk.BooleanVar(master=self, value=True)
        tkapp.filmtype_enabled_var = self.filmtype_enabled_var

        self._install_clipboard_shortcuts()

        self._build_header()
        self._build_tabs()
        self._apply_settings(self.settings, persist=False, apply_widgets=True)

        # Початкові дані
        self._refresh_categories()
        self._refresh_filmtype_checkboxes()

        if DEPENDENCY_WARNINGS:
            self.after(200, self._show_dependency_warnings)

        self.after(100, self._process_ui_queue)

    def _setup_ttk_styles(self) -> None:
        try:
            style = ttk.Style(self)
        except Exception:
            return

        try:
            current_theme = style.theme_use()
        except Exception:
            current_theme = ""

        if current_theme.lower() != "clam":
            try:
                style.theme_use("clam")
                logger.info("Switched ttk theme from %s to clam", current_theme)
            except Exception:
                pass

        theme_colors = self.theme_manager.colors or {}
        theme_fonts = self.theme_manager.fonts or {}
        base_bg = theme_colors.get("widget_fg", "#1b1d21")
        base_fg = theme_colors.get("text", "#f1f5f9")
        border = theme_colors.get("border", base_bg)
        highlight = theme_colors.get("accent", "#1f6aa5")
        header_bg = theme_colors.get("header_bg", highlight)
        header_text = theme_colors.get("header_text", "#ffffff")
        header_border = theme_colors.get("header_border", border)
        selection_bg = theme_colors.get("selection_bg", highlight)
        selection_text = theme_colors.get("selection_text", "#ffffff")
        scrollbar_track = theme_colors.get("scrollbar_track", base_bg)
        scrollbar_thumb = theme_colors.get("scrollbar_thumb", highlight)
        scrollbar_thumb_hover = theme_colors.get("scrollbar_thumb_hover", scrollbar_thumb)
        logger.info(
            "Applying ttk colors scrollbar_track=%s scrollbar_thumb=%s header_bg=%s",
            scrollbar_track,
            scrollbar_thumb,
            header_bg,
        )

        family = str(theme_fonts.get("family", "Segoe UI"))
        body_size = int(theme_fonts.get("base_size", 13))
        heading_size = int(theme_fonts.get("heading_size", 13))

        body_font = ctk.CTkFont(family=family, size=body_size)
        heading_font = ctk.CTkFont(family=family, size=heading_size, weight="bold")

        body_family = body_font.actual("family") or body_font.cget("family")
        heading_family = heading_font.actual("family") or heading_font.cget("family")
        body_size = body_font.cget("size")
        heading_size = heading_font.cget("size")

        def safe_style_configure(style_name: str, **kwargs: object) -> None:
            try:
                style.configure(style_name, **kwargs)
            except Exception:
                logger.exception("Не вдалося налаштувати ttk стиль %s", style_name)

        def safe_style_map(style_name: str, **kwargs: object) -> None:
            try:
                style.map(style_name, **kwargs)
            except Exception:
                logger.exception("Не вдалося налаштувати ttk map для %s", style_name)

        def safe_style_layout(style_name: str, layout: list[tuple[str, dict]]) -> None:
            try:
                style.layout(style_name, layout)
            except Exception:
                logger.exception("Не вдалося налаштувати ttk layout для %s", style_name)

        safe_style_configure(
            "Treeview",
            background=base_bg,
            fieldbackground=base_bg,
            foreground=base_fg,
            rowheight=26,
            font=(body_family, body_size),
            bordercolor=border,
            borderwidth=0,
            relief="flat",
            padding=0,
        )
        safe_style_configure(
            "Treeview.Heading",
            background=header_bg,
            foreground=header_text,
            font=(heading_family, heading_size, "bold"),
            relief="flat",
            borderwidth=1,
            bordercolor=header_border,
            padding=(6, 4),
        )
        safe_style_map(
            "Treeview",
            background=[("selected", selection_bg)],
            foreground=[("selected", selection_text)],
        )
        safe_style_map(
            "Treeview.Heading",
            background=[("active", header_bg)],
            foreground=[("active", header_text)],
            relief=[("active", "flat")],
        )
        safe_style_layout(
            "Treeview",
            [
                (
                    "Treeview.treearea",
                    {
                        "sticky": "nswe",
                    },
                )
            ],
        )
        safe_style_configure(
            "TScrollbar",
            troughcolor=scrollbar_track,
            background=scrollbar_thumb,
            bordercolor=scrollbar_track,
            arrowcolor=scrollbar_thumb,
            lightcolor=scrollbar_thumb,
            darkcolor=scrollbar_thumb,
        )
        safe_style_configure(
            "Vertical.TScrollbar",
            troughcolor=scrollbar_track,
            background=scrollbar_thumb,
            bordercolor=scrollbar_track,
            arrowcolor=scrollbar_thumb,
            lightcolor=scrollbar_thumb,
            darkcolor=scrollbar_thumb,
        )
        safe_style_configure(
            "Horizontal.TScrollbar",
            troughcolor=scrollbar_track,
            background=scrollbar_thumb,
            bordercolor=scrollbar_track,
            arrowcolor=scrollbar_thumb,
            lightcolor=scrollbar_thumb,
            darkcolor=scrollbar_thumb,
        )
        safe_style_map(
            "TScrollbar",
            background=[("active", scrollbar_thumb_hover), ("pressed", scrollbar_thumb_hover)],
        )
        safe_style_map(
            "Vertical.TScrollbar",
            background=[("active", scrollbar_thumb_hover), ("pressed", scrollbar_thumb_hover)],
        )
        safe_style_map(
            "Horizontal.TScrollbar",
            background=[("active", scrollbar_thumb_hover), ("pressed", scrollbar_thumb_hover)],
        )

    def _ensure_background_primitives(self) -> None:
        if not hasattr(self, "_progress_lock") or self._progress_lock is None:
            self._progress_lock = threading.Lock()
        if not hasattr(self, "_ui_event_queue") or self._ui_event_queue is None:
            self._ui_event_queue = queue.Queue()
        if not hasattr(self, "_ui_queue_job"):
            self._ui_queue_job = None
        if not hasattr(self, "_generation_task_running"):
            self._generation_task_running = False
        if not hasattr(self, "_active_generation_thread"):
            self._active_generation_thread = None
        if not hasattr(self, "_last_progress_update"):
            self._last_progress_update = 0.0

    def _process_ui_queue(self) -> None:
        self._ensure_background_primitives()
        if not hasattr(self, "_ui_event_queue"):
            return
        try:
            while True:
                callback = self._ui_event_queue.get_nowait()
                try:
                    callback()
                except Exception:
                    logger.exception("Не вдалося виконати відкладену дію інтерфейсу")
        except queue.Empty:
            pass
        try:
            self._ui_queue_job = self.after(60, self._process_ui_queue)
        except Exception:
            self._ui_queue_job = None

    def _call_in_ui_thread(self, func: Callable[..., None], *args, **kwargs) -> None:
        try:
            self._ensure_background_primitives()
        except Exception:
            pass

        after_fn = getattr(self, "after", None)
        if after_fn is None or not callable(after_fn):
            try:
                func(*args, **kwargs)
            except Exception:
                logger.exception("Помилка під час оновлення інтерфейсу")
            return

        queue_obj = getattr(self, "_ui_event_queue", None)
        if queue_obj is None:
            try:
                func(*args, **kwargs)
            except Exception:
                logger.exception("Помилка під час оновлення інтерфейсу")
            return

        def _wrapper() -> None:
            try:
                func(*args, **kwargs)
            except Exception:
                logger.exception("Помилка під час оновлення інтерфейсу")

        queue_obj.put(_wrapper)

    def _start_background_task(self, target: Callable[[], None], name: str = "background-task") -> threading.Thread:
        self._ensure_background_primitives()
        after_fn = getattr(self, "after", None)
        if after_fn is None or not callable(after_fn):
            target()
            return threading.current_thread()
        thread = threading.Thread(target=target, name=name, daemon=True)
        thread.start()
        return thread

    def _install_clipboard_shortcuts(self) -> None:
        sequences = [
            ("<Control-c>", "<<Copy>>"),
            ("<Control-v>", "<<Paste>>"),
            ("<Control-x>", "<<Cut>>"),
            ("<Control-a>", "<<SelectAll>>"),
        ]

        if sys.platform == "darwin":
            sequences.extend(
                [
                    ("<Command-c>", "<<Copy>>"),
                    ("<Command-v>", "<<Paste>>"),
                    ("<Command-x>", "<<Cut>>"),
                    ("<Command-a>", "<<SelectAll>>"),
                ]
            )

        unique = []
        seen = set()
        for sequence, virtual_event in sequences:
            key = (sequence, virtual_event)
            if key in seen:
                continue
            seen.add(key)
            unique.append((sequence, virtual_event))

        self._clipboard_shortcut_sequences = tuple(unique)
        self._clipboard_bound_widgets = set()

    def _resolve_clipboard_target(self, widget):
        if widget is None:
            return None
        if isinstance(widget, (tk.Entry, tk.Text)):
            return widget
        target = None
        if isinstance(widget, ctk.CTkEntry):
            target = getattr(widget, "_entry", None)
        elif isinstance(widget, ctk.CTkTextbox):
            target = getattr(widget, "_textbox", None)
            if target is None:
                target = getattr(widget, "_text_widget", None)
        if target is None and hasattr(widget, "event_generate") and hasattr(widget, "bind"):
            target = widget
        return target

    def _bind_clipboard_shortcuts(self, widget) -> None:
        sequences = getattr(self, "_clipboard_shortcut_sequences", ())
        if not sequences:
            return
        target = self._resolve_clipboard_target(widget)
        if target is None:
            return
        widget_id = str(target)
        bound_ids = getattr(self, "_clipboard_bound_widgets", set())
        if widget_id in bound_ids:
            return
        bound = False
        for sequence, virtual_event in sequences:
            try:
                target.bind(
                    sequence,
                    lambda event, ve=virtual_event, tw=target: self._handle_clipboard_shortcut(event, tw, ve),
                    add="+",
                )
            except Exception:
                continue
            bound = True
        if bound:
            bound_ids.add(widget_id)
            self._clipboard_bound_widgets = bound_ids

    def _bind_clipboard_context_menu(self, widget) -> None:
        target = self._resolve_clipboard_target(widget)
        if target is None:
            return

        menu = getattr(target, "_clipboard_context_menu", None)
        if not isinstance(menu, tk.Menu):
            menu = tk.Menu(target, tearoff=0)
            commands = (
                ("Cut", self._clipboard_cut),
                ("Copy", self._clipboard_copy),
                ("Paste", self._clipboard_paste),
                ("Select All", self._clipboard_select_all),
            )

            def _invoke_clipboard(action, fallback_target=target):
                resolved = self._resolve_clipboard_target(self.focus_get())
                destination = resolved or fallback_target
                if destination is None:
                    return None
                try:
                    destination.focus_set()
                except Exception:
                    pass
                try:
                    if action(destination):
                        return "break"
                except Exception:
                    return None
                return None

            for label, action in commands:
                menu.add_command(
                    label=label,
                    command=lambda fn=action: _invoke_clipboard(fn),
                )

            setattr(target, "_clipboard_context_menu", menu)

        def _show_menu(event, ctx_menu=menu, fallback_target=target):
            active = self._resolve_clipboard_target(event.widget) or fallback_target
            try:
                active.focus_set()
            except Exception:
                pass
            try:
                ctx_menu.tk_popup(event.x_root, event.y_root)
            finally:
                ctx_menu.grab_release()
            return "break"

        sequences = {"<Button-2>", "<Button-3>", "<Shift-F10>"}
        if sys.platform == "darwin":
            sequences.add("<Control-Button-1>")

        def _bind_sequences(target_widget, seen_attr: str):
            if target_widget is None:
                return
            seen = getattr(target_widget, seen_attr, set())
            if not isinstance(seen, set):
                seen = set()
            for sequence in sequences:
                if sequence in seen:
                    continue
                try:
                    target_widget.bind(sequence, _show_menu, add="+")
                except Exception:
                    continue
                seen.add(sequence)
            setattr(target_widget, seen_attr, seen)

        _bind_sequences(target, "_clipboard_context_sequences")
        if widget is not target:
            _bind_sequences(widget, "_clipboard_context_sequences")
    def _clipboard_get_selection_text(self, widget):
        if widget is None:
            return None
        try:
            return widget.selection_get()
        except Exception:
            pass
        try:
            start = widget.index("sel.first")
            end = widget.index("sel.last")
            return widget.get(start, end)
        except Exception:
            return None

    def _clipboard_delete_selection(self, widget) -> bool:
        if widget is None:
            return False
        try:
            start = widget.index("sel.first")
            end = widget.index("sel.last")
        except Exception:
            return False
        try:
            widget.delete(start, end)
        except Exception:
            return False
        return True

    def _clipboard_copy(self, widget) -> bool:
        if widget is None:
            return False
        try:
            widget.event_generate("<<Copy>>")
            return True
        except Exception:
            pass
        text = self._clipboard_get_selection_text(widget)
        if text is None:
            return False
        try:
            widget.clipboard_clear()
            widget.clipboard_append(text)
        except Exception:
            return False
        return True

    def _clipboard_cut(self, widget) -> bool:
        if widget is None:
            return False
        try:
            widget.event_generate("<<Cut>>")
            return True
        except Exception:
            pass
        text = self._clipboard_get_selection_text(widget)
        if text is None:
            return False
        try:
            widget.clipboard_clear()
            widget.clipboard_append(text)
        except Exception:
            return False
        return self._clipboard_delete_selection(widget)

    def _clipboard_paste(self, widget) -> bool:
        if widget is None:
            return False
        try:
            widget.event_generate("<<Paste>>")
            return True
        except Exception:
            pass
        try:
            data = widget.clipboard_get()
        except Exception:
            return False
        if data is None:
            data = ""
        self._clipboard_delete_selection(widget)
        try:
            widget.insert(tk.INSERT, data)
        except Exception:
            return False
        return True

    def _clipboard_select_all(self, widget) -> bool:
        if widget is None:
            return False
        try:
            widget.event_generate("<<SelectAll>>")
            return True
        except Exception:
            pass
        try:
            if isinstance(widget, tk.Entry):
                widget.select_range(0, tk.END)
                widget.icursor(tk.END)
            else:
                widget.tag_add("sel", "1.0", "end-1c")
                widget.mark_set("insert", "end-1c")
                widget.see("insert")
        except Exception:
            return False
        return True

    def _handle_clipboard_shortcut(self, event, widget, virtual_event: str):
        target = self._resolve_clipboard_target(widget) or widget
        if target is None:
            return None
        try:
            target.event_generate(virtual_event)
        except Exception:
            return None
        return "break"

    def _sync_templates_with_catalog(self):
        categories = [name.strip() for _cid, name in get_categories() if isinstance(name, str) and name.strip()]
        descriptions = self.templates.setdefault("descriptions", {})
        changed_templates = False
        for name in categories:
            if name not in descriptions:
                descriptions[name] = {}
                changed_templates = True
        if changed_templates:
            save_templates(self.templates)

        changed_title_tags = False
        for name in categories:
            if self._ensure_title_tags_category(name):
                changed_title_tags = True

        for item in self.templates.get("film_types", []):
            fname = item.get("name")
            if isinstance(fname, str) and fname:
                if self._ensure_title_tags_film(fname):
                    changed_title_tags = True

        if changed_title_tags:
            save_title_tags_templates(self.title_tags_templates)

    def _ensure_title_tags_category(self, category_name: str) -> bool:
        if not category_name:
            return False
        by_category = self.title_tags_templates.setdefault("by_category", {})
        changed = False
        cat_entry = by_category.get(category_name)
        if not isinstance(cat_entry, dict):
            by_category[category_name] = {"default": {}, "by_film": {}}
            return True
        if "default" not in cat_entry or not isinstance(cat_entry["default"], dict):
            cat_entry["default"] = {}
            changed = True
        if "by_film" not in cat_entry or not isinstance(cat_entry["by_film"], dict):
            cat_entry["by_film"] = {}
            changed = True
        return changed

    def _ensure_title_tags_film(self, film_name: str) -> bool:
        if not film_name:
            return False
        by_film = self.title_tags_templates.setdefault("by_film", {})
        block = by_film.get(film_name)
        if not isinstance(block, dict):
            by_film[film_name] = {}
            return True
        return False

    def _rename_category_templates(self, old_name: str, new_name: str):
        if not old_name or not new_name or old_name == new_name:
            return
        descriptions = self.templates.setdefault("descriptions", {})
        changed_templates = False
        old_block = descriptions.pop(old_name, None)
        if isinstance(old_block, dict):
            target = descriptions.get(new_name)
            if isinstance(target, dict):
                for key, value in old_block.items():
                    if key not in target:
                        target[key] = value
            else:
                descriptions[new_name] = old_block
            changed_templates = True

        by_category = self.title_tags_templates.setdefault("by_category", {})
        changed_title_tags = False
        old_cat_block = by_category.pop(old_name, None)
        if isinstance(old_cat_block, dict):
            target_block = by_category.get(new_name)
            if isinstance(target_block, dict):
                if isinstance(old_cat_block.get("default"), dict) and not isinstance(target_block.get("default"), dict):
                    target_block["default"] = old_cat_block["default"]
                src_films = old_cat_block.get("by_film")
                if isinstance(src_films, dict):
                    dst_films = target_block.setdefault("by_film", {})
                    for film_name, block in src_films.items():
                        if film_name not in dst_films:
                            dst_films[film_name] = block
            else:
                by_category[new_name] = old_cat_block
            changed_title_tags = True

        if changed_templates:
            save_templates(self.templates)
        if changed_title_tags or self._ensure_title_tags_category(new_name):
            save_title_tags_templates(self.title_tags_templates)

    def _delete_category_templates(self, category_name: str):
        if not category_name:
            return
        descriptions = self.templates.get("descriptions", {})
        changed_templates = False
        if category_name in descriptions:
            descriptions.pop(category_name, None)
            changed_templates = True

        by_category = self.title_tags_templates.get("by_category", {})
        changed_title_tags = False
        if isinstance(by_category, dict) and category_name in by_category:
            by_category.pop(category_name, None)
            changed_title_tags = True

        if changed_templates:
            save_templates(self.templates)
        if changed_title_tags:
            save_title_tags_templates(self.title_tags_templates)

    def _rename_film_type(self, old_name: str, new_name: str):
        if not old_name or not new_name or old_name == new_name:
            return
        changed_templates = False
        for desc in self.templates.get("descriptions", {}).values():
            if isinstance(desc, dict) and old_name in desc:
                if new_name not in desc:
                    desc[new_name] = desc.pop(old_name)
                else:
                    desc.pop(old_name, None)
                changed_templates = True

        by_film = self.title_tags_templates.setdefault("by_film", {})
        changed_title_tags = False
        if isinstance(by_film, dict) and old_name in by_film:
            block = by_film.pop(old_name)
            by_film[new_name] = block
            changed_title_tags = True

        for cat_block in self.title_tags_templates.setdefault("by_category", {}).values():
            if not isinstance(cat_block, dict):
                continue
            films_map = cat_block.get("by_film")
            if isinstance(films_map, dict) and old_name in films_map:
                if new_name not in films_map:
                    films_map[new_name] = films_map.pop(old_name)
                else:
                    films_map.pop(old_name, None)
                changed_title_tags = True

        if changed_templates:
            save_templates(self.templates)
        if changed_title_tags:
            save_title_tags_templates(self.title_tags_templates)

    def _remove_film_type_templates(self, film_name: str):
        if not film_name:
            return
        changed_templates = False
        for desc in self.templates.get("descriptions", {}).values():
            if isinstance(desc, dict) and film_name in desc:
                desc.pop(film_name, None)
                changed_templates = True

        by_film = self.title_tags_templates.get("by_film", {})
        changed_title_tags = False
        if isinstance(by_film, dict) and film_name in by_film:
            by_film.pop(film_name, None)
            changed_title_tags = True

        for cat_block in self.title_tags_templates.get("by_category", {}).values():
            if not isinstance(cat_block, dict):
                continue
            films_map = cat_block.get("by_film")
            if isinstance(films_map, dict) and film_name in films_map:
                films_map.pop(film_name, None)
                changed_title_tags = True

        if changed_templates:
            save_templates(self.templates)
        if changed_title_tags:
            save_title_tags_templates(self.title_tags_templates)

    # -------- верхній бар
    def _build_header(self):
        top = ctk.CTkFrame(self)
        top.pack(fill="x", padx=10, pady=(10, 0))
        self.header_frame = top
        self.file_menu_button = ctk.CTkButton(
            top,
            text="Файл",
            width=80,
            command=self._open_file_menu,
        )
        self.file_menu_button.pack(side="left")
        self.theme_manager.register(top, "surface")
        self.theme_manager.register(self.file_menu_button, "menu_button")

    def _open_file_menu(self):
        if self._file_menu is None:
            menu = tk.Menu(self, tearoff=0)
            menu.add_command(label="Налаштування...", command=self._open_settings_dialog)
            menu.add_separator()
            menu.add_command(label="Відкрити папку даних", command=self._open_data_folder)
            menu.add_command(label="Відкрити папку логів", command=self._open_logs_folder)
            menu.add_separator()
            menu.add_command(label="Про програму", command=self._show_about)
            menu.add_command(label="Вихід", command=self._handle_exit)
            self._file_menu = menu
        if self.file_menu_button is None:
            return
        x = self.file_menu_button.winfo_rootx()
        y = self.file_menu_button.winfo_rooty() + self.file_menu_button.winfo_height()
        try:
            self._file_menu.tk_popup(x, y)
        finally:
            self._file_menu.grab_release()

    def _open_settings_dialog(self):
        SettingsDialog(self, self.settings, on_apply=self._on_settings_applied)

    def _on_settings_applied(self, settings: Dict[str, object]) -> None:
        validated = validate_settings(settings)
        self._apply_settings(validated, persist=True, apply_widgets=True)

    def _apply_settings(self, settings: Dict[str, object], *, persist: bool, apply_widgets: bool) -> None:
        self.settings = settings
        self.theme_manager.apply(settings, apply_widgets=apply_widgets)
        self._setup_ttk_styles()
        if persist:
            save_settings(settings)
        if self.out_folder_var is not None:
            export_folder = settings.get("export_folder")
            if isinstance(export_folder, str) and export_folder:
                self.out_folder_var.set(export_folder)

    def _handle_exit(self):
        self.destroy()

    def _show_about(self):
        data_dir = get_data_dir()
        messagebox.showinfo(
            APP_TITLE,
            f"{APP_TITLE}\nКаталог даних: {data_dir}",
        )

    def _open_data_folder(self):
        self._open_path(get_data_dir())

    def _open_logs_folder(self):
        self._open_path(get_logs_dir())

    def _open_path(self, path: Path):
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", str(path)], check=False)
            else:
                subprocess.run(["xdg-open", str(path)], check=False)
        except Exception:
            logger.exception("Не вдалося відкрити шлях %s", path)
            show_error("Не вдалося відкрити папку.")

    # -------- вкладки
    def _build_tabs(self):
        tabs = ctk.CTkTabview(self, width=1040, height=600)
        tabs.pack(fill="both", expand=True, padx=10, pady=10)
        self.tabs = tabs
        self.theme_manager.register(tabs, "tabview")

        self.tab_catalog   = tabs.add("Каталог")
        self.tab_templates = tabs.add("Шаблони")
        self.tab_parameters = tabs.add("Параметри")
        self.tab_export    = tabs.add("Експорт")
        self.tab_files     = tabs.add("Файли")
        self.tab_generate  = tabs.add("Генерація")

        self._build_tab_catalog()
        self._build_tab_templates()
        self._build_tab_parameters()
        self._build_tab_export()
        self._build_tab_files()
        self._build_tab_generate()

    def _show_dependency_warnings(self):
        for warning in DEPENDENCY_WARNINGS:
            messagebox.showwarning(APP_TITLE, warning)
        DEPENDENCY_WARNINGS.clear()

    def _film_type_names(self):
        return [item.get("name") for item in self.templates.get("film_types", []) if item.get("name")]

    def _film_type_menu_items(self):
        items = [(FILM_TYPE_DEFAULT_LABEL, "default")]
        for name in self._film_type_names():
            items.append((name, name))
        return items

    def _template_language_items(self):
        items = [(TEMPLATE_LANGUAGE_DEFAULT_LABEL, None)]
        languages = self.templates.get("template_languages", [])
        if isinstance(languages, list):
            seen = set()
            for item in languages:
                code = item.get("code") if isinstance(item, dict) else None
                label = item.get("label") if isinstance(item, dict) else None
                if not isinstance(code, str):
                    continue
                stripped = code.strip()
                if not stripped or stripped.lower() in seen:
                    continue
                seen.add(stripped.lower())
                if not isinstance(label, str) or not label.strip():
                    label = stripped
                items.append((label.strip(), stripped))
        return items

    def _template_language_codes(self):
        return [code for label, code in self._template_language_items() if code]

    def _language_label_for_code(self, code: str) -> str:
        mapping = getattr(self, "_template_language_code_to_label", {})
        if isinstance(mapping, dict) and code in mapping:
            return mapping[code]
        for label, value in self._template_language_items():
            if value == code:
                return label
        return code

    def _template_category_items(self):
        names = set()
        for _cid, name in get_categories():
            if isinstance(name, str):
                stripped = name.strip()
                if stripped:
                    names.add(stripped)
        for name in self.templates.get("descriptions", {}).keys():
            if isinstance(name, str):
                stripped = name.strip()
                if stripped and stripped != GLOBAL_DESCRIPTION_KEY:
                    names.add(stripped)
        items = [(CATEGORY_SCOPE_DEFAULT_LABEL, None)]
        for name in sorted(names):
            items.append((name, name))
        return items

    def _refresh_template_selectors(self):
        if not hasattr(self, "template_category_menu") or not hasattr(self, "template_film_menu"):
            return

        category_items = self._template_category_items()
        if not category_items:
            category_items = [(CATEGORY_SCOPE_DEFAULT_LABEL, None)]
        film_items = self._film_type_menu_items()
        if not film_items:
            film_items = [(FILM_TYPE_DEFAULT_LABEL, "default")]

        self._template_category_label_to_key = {label: key for label, key in category_items}
        self._template_category_key_to_label = {key: label for label, key in category_items}
        self._template_film_label_to_key = {label: key for label, key in film_items}
        self._template_film_key_to_label = {key: label for label, key in film_items}

        if hasattr(self, "template_category_menu"):
            self.template_category_menu.configure(values=[label for label, _ in category_items])
        if hasattr(self, "template_film_menu"):
            self.template_film_menu.configure(values=[label for label, _ in film_items])

        current_cat = self._current_template_category
        if current_cat not in self._template_category_key_to_label:
            current_cat = category_items[0][1]
            if current_cat is None and len(category_items) > 1:
                current_cat = category_items[1][1]
        self._current_template_category = current_cat
        category_label = self._template_category_key_to_label.get(current_cat, CATEGORY_SCOPE_DEFAULT_LABEL)
        self.template_category_var.set(category_label)
        self.template_category_menu.set(category_label)

        current_film = self._selected_film_type_key()
        if current_film not in self._template_film_key_to_label:
            current_film = film_items[0][1]
        self._current_film_type_key = current_film
        film_label = self._template_film_key_to_label.get(current_film, FILM_TYPE_DEFAULT_LABEL)
        self.template_film_var.set(film_label)
        self.template_film_menu.set(film_label)

        if hasattr(self, "template_language_menu") and hasattr(self, "template_language_var"):
            language_items = self._template_language_items()
            self._template_language_label_to_code = {label: code for label, code in language_items}
            self._template_language_code_to_label = {code: label for label, code in language_items}
            self.template_language_menu.configure(values=[label for label, _ in language_items])
            current_lang = self._current_template_language
            if current_lang not in self._template_language_code_to_label:
                current_lang = language_items[0][1]
            self._current_template_language = current_lang
            language_label = self._template_language_code_to_label.get(current_lang, TEMPLATE_LANGUAGE_DEFAULT_LABEL)
            self.template_language_var.set(language_label)
            self.template_language_menu.set(language_label)

        self._on_template_scope_change()

    def _set_title_tags_block(
        self,
        category_key: Optional[str],
        film_key: str,
        language_code: Optional[str],
        title_value: str,
        tags_value: str,
    ):
        fallback_block = _title_tags_block(
            self.templates.get("title_template", DEFAULT_TEMPLATES["title_template"]),
            self.templates.get("tags_template", DEFAULT_TEMPLATES["tags_template"]),
        )

        def _update_block(container: dict, key: str) -> None:
            existing = container.get(key)
            normalized = _normalize_title_tags_block(existing, fallback_block)
            normalized["title_template"] = _set_language_template_value(
                normalized.get("title_template"), language_code, title_value, fallback_block["title_template"].get("default", "")
            )
            normalized["tags_template"] = _set_language_template_value(
                normalized.get("tags_template"), language_code, tags_value, fallback_block["tags_template"].get("default", "")
            )
            container[key] = normalized

        root_default = self.title_tags_templates.get("default")
        if not isinstance(root_default, dict):
            self.title_tags_templates["default"] = deepcopy(fallback_block)

        if category_key:
            self._ensure_title_tags_category(category_key)
            by_category = self.title_tags_templates.setdefault("by_category", {})
            cat_entry = by_category.setdefault(category_key, {"default": {}, "by_film": {}})
            if film_key == "default":
                _update_block(cat_entry, "default")
            else:
                films = cat_entry.setdefault("by_film", {})
                _update_block(films, film_key)
        else:
            if film_key == "default":
                _update_block(self.title_tags_templates, "default")
            else:
                films = self.title_tags_templates.setdefault("by_film", {})
                _update_block(films, film_key)

    def _selected_film_type_key(self) -> str:
        key = getattr(self, "_current_film_type_key", None)
        if key in (None, ""):
            self._current_film_type_key = "default"
            return "default"
        if key == "default" or key in set(self._film_type_names()):
            return key
        self._current_film_type_key = "default"
        return "default"

    # -------- Каталог (перший дизайн)
    def _build_tab_catalog(self):
        # Ліва колона: Категорії + Бренди
        left = ctk.CTkFrame(self.tab_catalog)
        left.pack(side="left", fill="both", expand=True, padx=(0,10), pady=10)
        self.theme_manager.register(left, "surface")

        # Категорії
        cat_label = ctk.CTkLabel(left, text="Категорія")
        cat_label.pack(anchor="w", padx=10, pady=(8,0))
        self.theme_manager.register(cat_label, "label")
        cat_frame = ctk.CTkFrame(left)
        cat_frame.pack(fill="x", padx=10, pady=5)
        self.theme_manager.register(cat_frame, "surface")
        self.cat_tree = ttk.Treeview(cat_frame, columns=("name",), show="headings", height=6)
        self.cat_tree.heading("name", text="Назва")
        self.cat_tree.column("name", width=260, anchor="w")
        self.cat_tree.pack(side="left", fill="x", expand=True)
        cat_scroll = ttk.Scrollbar(cat_frame, orient="vertical", command=self.cat_tree.yview)
        cat_scroll.pack(side="right", fill="y"); self.cat_tree.configure(yscrollcommand=cat_scroll.set)
        self.cat_tree.bind("<<TreeviewSelect>>", self._on_category_select)
        self.cat_tree.bind("<Button-1>", lambda e: self._handle_tree_click(e, "cat", self.cat_tree), add="+")
        self.cat_tree.bind("<Delete>", lambda e: self._handle_tree_delete("cat"))

        cat_ctrl = ctk.CTkFrame(left)
        cat_ctrl.pack(fill="x", padx=10, pady=(0,10))
        self.theme_manager.register(cat_ctrl, "surface")
        self.cat_entry = ctk.CTkEntry(cat_ctrl, placeholder_text="Назва категорії")
        self.cat_entry.pack(side="left", fill="x", expand=True, padx=(0,5))
        self._bind_clipboard_shortcuts(self.cat_entry)
        self.theme_manager.register(self.cat_entry, "widget")
        self.cat_add_button = ctk.CTkButton(cat_ctrl, text="Додати", command=self._cat_add, width=90)
        self.cat_add_button.pack(side="left", padx=3)
        self.theme_manager.register(self.cat_add_button, "accent_button")
        self.cat_rename_button = ctk.CTkButton(cat_ctrl, text="Перейменувати", command=self._cat_rename, width=120)
        self.cat_rename_button.pack(side="left", padx=3)
        self.theme_manager.register(self.cat_rename_button, "accent_button")
        self.cat_delete_button = ctk.CTkButton(cat_ctrl, text="Видалити", command=self._cat_delete, width=90)
        self.cat_delete_button.pack(side="left", padx=3)
        self.theme_manager.register(self.cat_delete_button, "danger_button")

        # Бренди
        brand_label = ctk.CTkLabel(left, text="Бренд")
        brand_label.pack(anchor="w", padx=10, pady=(8,0))
        self.theme_manager.register(brand_label, "label")
        brand_frame = ctk.CTkFrame(left)
        brand_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.theme_manager.register(brand_frame, "surface")
        self.brand_tree = ttk.Treeview(brand_frame, columns=("name",), show="headings", height=11)
        self.brand_tree.heading("name", text="Назва")
        self.brand_tree.column("name", width=260, anchor="w")
        self.brand_tree.pack(side="left", fill="both", expand=True)
        brand_scroll = ttk.Scrollbar(brand_frame, orient="vertical", command=self.brand_tree.yview)
        brand_scroll.pack(side="right", fill="y"); self.brand_tree.configure(yscrollcommand=brand_scroll.set)
        self.brand_tree.bind("<<TreeviewSelect>>", self._on_brand_select)
        self.brand_tree.bind("<Button-1>", lambda e: self._handle_tree_click(e, "brand", self.brand_tree), add="+")
        self.brand_tree.bind("<Delete>", lambda e: self._handle_tree_delete("brand"))

        brand_ctrl = ctk.CTkFrame(left)
        brand_ctrl.pack(fill="x", padx=10, pady=(0,10))
        self.theme_manager.register(brand_ctrl, "surface")
        self.brand_entry = ctk.CTkEntry(brand_ctrl, placeholder_text="Назва бренду")
        self.brand_entry.pack(side="left", fill="x", expand=True, padx=(0,5))
        self._bind_clipboard_shortcuts(self.brand_entry)
        self.theme_manager.register(self.brand_entry, "widget")
        self.brand_add_button = ctk.CTkButton(brand_ctrl, text="Додати", command=self._brand_add, width=90)
        self.brand_add_button.pack(side="left", padx=3)
        self.theme_manager.register(self.brand_add_button, "accent_button")
        self.brand_rename_button = ctk.CTkButton(brand_ctrl, text="Перейменувати", command=self._brand_rename, width=120)
        self.brand_rename_button.pack(side="left", padx=3)
        self.theme_manager.register(self.brand_rename_button, "accent_button")
        self.brand_delete_button = ctk.CTkButton(brand_ctrl, text="Видалити", command=self._brand_delete, width=90)
        self.brand_delete_button.pack(side="left", padx=3)
        self.theme_manager.register(self.brand_delete_button, "danger_button")

        # Права колона: Моделі
        right = ctk.CTkFrame(self.tab_catalog)
        right.pack(side="left", fill="both", expand=True, padx=(10,0), pady=10)
        self.theme_manager.register(right, "surface")

        model_label = ctk.CTkLabel(right, text="Модель")
        model_label.pack(anchor="w", padx=10, pady=(8,0))
        self.theme_manager.register(model_label, "label")
        model_frame = ctk.CTkFrame(right)
        model_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.theme_manager.register(model_frame, "surface")
        self.model_tree = ttk.Treeview(model_frame, columns=("name",), show="headings", height=22)
        self.model_tree.heading("name", text="Назва")
        self.model_tree.column("name", width=360, anchor="w")
        self.model_tree.pack(side="left", fill="both", expand=True)
        model_scroll = ttk.Scrollbar(model_frame, orient="vertical", command=self.model_tree.yview)
        model_scroll.pack(side="right", fill="y"); self.model_tree.configure(yscrollcommand=model_scroll.set)
        self.model_tree.bind("<Button-1>", lambda e: self._handle_tree_click(e, "model", self.model_tree), add="+")
        self.model_tree.bind("<Double-1>", self._on_model_double_click, add="+")
        self.model_tree.bind("<Delete>", lambda e: self._handle_tree_delete("model"))

        model_ctrl = ctk.CTkFrame(right)
        model_ctrl.pack(fill="x", padx=10, pady=(0,10))
        self.theme_manager.register(model_ctrl, "surface")
        self.model_entry = ctk.CTkEntry(model_ctrl, placeholder_text="Назва моделі")
        self.model_entry.pack(side="left", fill="x", expand=True, padx=(0,5))
        self._bind_clipboard_shortcuts(self.model_entry)
        self.theme_manager.register(self.model_entry, "widget")
        self.model_add_button = ctk.CTkButton(model_ctrl, text="Додати", command=self._model_add, width=90)
        self.model_add_button.pack(side="left", padx=3)
        self.theme_manager.register(self.model_add_button, "accent_button")
        self.model_rename_button = ctk.CTkButton(model_ctrl, text="Перейменувати", command=self._model_rename, width=120)
        self.model_rename_button.pack(side="left", padx=3)
        self.theme_manager.register(self.model_rename_button, "accent_button")
        self.model_delete_button = ctk.CTkButton(model_ctrl, text="Видалити", command=self._model_delete, width=90)
        self.model_delete_button.pack(side="left", padx=3)
        self.theme_manager.register(self.model_delete_button, "danger_button")
        self.model_specs_button = ctk.CTkButton(model_ctrl, text="Характеристики", command=self._open_specs, width=140)
        self.model_specs_button.pack(side="left", padx=6)
        self.theme_manager.register(self.model_specs_button, "accent_button")

    def _on_model_double_click(self, event):
        row = self.model_tree.identify_row(event.y)
        if not row:
            return
        self.model_tree.selection_set(row)
        self._open_specs()

    def _handle_tree_click(self, event, kind, tree):
        row = tree.identify_row(event.y)
        now = time.time()
        if not row:
            self._rename_clicks[kind] = (None, now)
            return
        last_row, last_time = self._rename_clicks.get(kind, (None, 0.0))
        self._rename_clicks[kind] = (row, now)
        delay = now - last_time
        if row == last_row and self._rename_delay_min <= delay <= self._rename_delay_max:
            self.after(0, lambda: self._start_tree_rename(kind, tree, row))

    def _handle_tree_delete(self, kind):
        if kind == "cat":
            self._cat_delete()
        elif kind == "brand":
            self._brand_delete()
        else:
            self._model_delete()
        return "break"

    def _start_tree_rename(self, kind, tree, iid):
        if not tree.exists(iid):
            return
        values = tree.item(iid, "values")
        if not values:
            return
        original = values[0]
        bbox = tree.bbox(iid, column="#1")
        if not bbox:
            return
        if self._rename_entry is not None:
            self._finish_inline_rename(save=False)
        entry = create_inline_entry(tree, original, theme_colors=self.theme_manager.colors)
        x, y, width, height = bbox
        entry.place(x=x, y=y, width=width, height=height)
        self._rename_entry = entry
        self._rename_entry_meta = (kind, tree, iid, original)
        entry.bind("<Return>", lambda _e: self._finish_inline_rename(save=True))
        entry.bind("<KP_Enter>", lambda _e: self._finish_inline_rename(save=True))
        entry.bind("<Escape>", lambda _e: self._finish_inline_rename(save=False))
        entry.bind("<FocusOut>", lambda _e: self._finish_inline_rename(save=True))

    def _finish_inline_rename(self, save: bool):
        if not self._rename_entry or not self._rename_entry_meta:
            return
        entry = self._rename_entry
        kind, tree, iid, original = self._rename_entry_meta
        self._rename_entry = None
        self._rename_entry_meta = None
        new_value = entry.get().strip()
        entry.destroy()
        if not save:
            self._restore_tree_selection(kind, iid)
            return
        if not new_value:
            if kind == "cat":
                show_error("Назва категорії не може бути порожньою.")
            elif kind == "brand":
                show_error("Назва бренду не може бути порожньою.")
            else:
                show_error("Назва моделі не може бути порожньою.")
            self._restore_tree_selection(kind, iid)
            return
        if new_value == original:
            self._restore_tree_selection(kind, iid)
            return
        self._apply_tree_rename(kind, iid, new_value)

    def _apply_tree_rename(self, kind, iid, new_value):
        if kind == "cat":
            cat_id = int(iid.split("_")[1])
            result = rename_category(cat_id, new_value)
            if result is not True:
                if isinstance(result, sqlite3.IntegrityError):
                    show_error("Категорія з такою назвою вже існує.")
                else:
                    show_error("Не вдалося перейменувати категорію.")
            self._refresh_categories()
            self.after(10, lambda: self._restore_tree_selection("cat", f"cat_{cat_id}"))
        elif kind == "brand":
            brand_id = int(iid.split("_")[1])
            result = rename_brand(brand_id, new_value)
            if result is not True:
                if isinstance(result, sqlite3.IntegrityError):
                    show_error("Бренд з такою назвою вже існує.")
                else:
                    show_error("Не вдалося перейменувати бренд.")
            if self.current_category_id:
                self._refresh_brands(self.current_category_id)
            else:
                self._refresh_brands(None)
            self._reload_gen_tree()
            self.after(10, lambda: self._restore_tree_selection("brand", f"brand_{brand_id}"))
        else:
            model_id = int(iid.split("_")[1])
            result = rename_model(model_id, new_value)
            if result is not True:
                if isinstance(result, sqlite3.IntegrityError):
                    show_error("Модель з такою назвою вже існує.")
                else:
                    show_error("Не вдалося перейменувати модель.")
            if self.current_brand_id:
                self._refresh_models(self.current_brand_id)
            else:
                self._refresh_models(None)
            self._reload_gen_tree()
            self.after(10, lambda: self._restore_tree_selection("model", f"model_{model_id}"))

    def _restore_tree_selection(self, kind, iid):
        tree = {
            "cat": getattr(self, "cat_tree", None),
            "brand": getattr(self, "brand_tree", None),
            "model": getattr(self, "model_tree", None),
        }.get(kind)
        if tree is None:
            return
        def _select():
            if not tree.exists(iid):
                return
            tree.selection_set(iid)
            tree.focus(iid)
            tree.see(iid)
            if kind == "cat":
                self._on_category_select()
            elif kind == "brand":
                self._on_brand_select()
        self.after(10, _select)

    # ---- catalog actions
    def _refresh_categories(self):
        self.cat_tree.delete(*self.cat_tree.get_children())
        for cid, name in get_categories():
            self.cat_tree.insert("", "end", iid=f"cat_{cid}", values=(name,))
        self.current_category_id = None
        self._refresh_brands(None)
        self._refresh_models(None)
        self._reload_gen_tree()
        self._sync_templates_with_catalog()
        self._refresh_template_selectors()

    def _refresh_brands(self, category_id):
        self.brand_tree.delete(*self.brand_tree.get_children())
        if category_id:
            for bid, name in get_brands(category_id):
                self.brand_tree.insert("", "end", iid=f"brand_{bid}", values=(name,))
        self.current_brand_id = None

    def _refresh_models(self, brand_id):
        self.model_tree.delete(*self.model_tree.get_children())
        if brand_id:
            for mid, name in get_models(brand_id):
                self.model_tree.insert("", "end", iid=f"model_{mid}", values=(name,))

    def _on_category_select(self, _evt=None):
        sel = self.cat_tree.selection()
        if not sel:
            self.current_category_id = None
            self._refresh_brands(None); self._refresh_models(None)
            return
        self.current_category_id = int(sel[0].split("_")[1])
        self._refresh_brands(self.current_category_id)
        self._refresh_models(None)

    def _on_brand_select(self, _evt=None):
        sel = self.brand_tree.selection()
        if not sel:
            self.current_brand_id = None
            self._refresh_models(None)
            return
        self.current_brand_id = int(sel[0].split("_")[1])
        self._refresh_models(self.current_brand_id)

    def _cat_add(self):
        name = self.cat_entry.get().strip()
        if not name: return show_error("Введіть назву категорії.")
        add_category(name); self.cat_entry.delete(0, tk.END); self._refresh_categories()

    def _cat_rename(self):
        if not self.current_category_id: return show_error("Виберіть категорію.")
        name = self.cat_entry.get().strip()
        if not name: return show_error("Введіть нову назву категорії.")
        sel = self.cat_tree.selection()
        old_name = None
        if sel:
            old_val = self.cat_tree.item(sel[0], "values")
            if old_val:
                old_name = (old_val[0] or "").strip()
        cat_id = self.current_category_id
        result = rename_category(cat_id, name)
        if result is not True:
            if isinstance(result, sqlite3.IntegrityError):
                show_error("Категорія з такою назвою вже існує.")
            else:
                show_error("Не вдалося перейменувати категорію.")
        else:
            if old_name:
                self._rename_category_templates(old_name, name)
            if self._current_template_category == old_name:
                self._current_template_category = name
            if self._current_desc_category == old_name:
                self._current_desc_category = name
        self._refresh_categories()
        if cat_id:
            self.after(10, lambda: self._restore_tree_selection("cat", f"cat_{cat_id}"))

    def _cat_delete(self):
        selection = list(self.cat_tree.selection())
        if not selection:
            return show_error("Виберіть категорію.")
        prompt = (
            "Видалити вибрану категорію та всі її бренди/моделі?"
            if len(selection) == 1
            else "Видалити вибрані категорії та всі їх бренди/моделі?"
        )
        if not messagebox.askyesno("Підтвердження", prompt):
            return
        ids = [int(iid.split("_")[1]) for iid in selection]
        id_to_name = {}
        for iid in selection:
            parts = iid.split("_")
            if len(parts) == 2:
                try:
                    cid = int(parts[1])
                except ValueError:
                    continue
                values = self.cat_tree.item(iid, "values")
                if values:
                    id_to_name[cid] = (values[0] or "").strip()
        for cat_id in ids:
            delete_category(cat_id)
            cat_name = id_to_name.get(cat_id)
            if cat_name:
                self._delete_category_templates(cat_name)
                if self._current_template_category == cat_name:
                    self._current_template_category = None
                if self._current_desc_category == cat_name:
                    self._current_desc_category = None
        self._refresh_categories()

    def _brand_add(self):
        if not self.current_category_id: return show_error("Спочатку виберіть категорію.")
        raw = self.brand_entry.get()
        names = split_catalog_input(raw)
        if not names: return show_error("Введіть назву бренду (через кому для декількох).")
        for name in names:
            add_brand(self.current_category_id, name)
        self.brand_entry.delete(0, tk.END)
        self._refresh_brands(self.current_category_id)
        self._reload_gen_tree()

    def _brand_rename(self):
        if not self.current_brand_id: return show_error("Виберіть бренд.")
        name = self.brand_entry.get().strip()
        if not name: return show_error("Введіть нову назву бренду.")
        brand_id = self.current_brand_id
        result = rename_brand(brand_id, name)
        if result is not True:
            if isinstance(result, sqlite3.IntegrityError):
                show_error("Бренд з такою назвою вже існує.")
            else:
                show_error("Не вдалося перейменувати бренд.")
        self._refresh_brands(self.current_category_id)
        if brand_id:
            self.after(10, lambda: self._restore_tree_selection("brand", f"brand_{brand_id}"))
        self._reload_gen_tree()

    def _brand_delete(self):
        selection = list(self.brand_tree.selection())
        if not selection:
            return show_error("Виберіть бренд.")
        prompt = (
            "Видалити вибраний бренд та всі його моделі?"
            if len(selection) == 1
            else "Видалити вибрані бренди та всі їх моделі?"
        )
        if not messagebox.askyesno("Підтвердження", prompt):
            return
        ids = [int(iid.split("_")[1]) for iid in selection]
        for brand_id in ids:
            delete_brand(brand_id)
        self._refresh_brands(self.current_category_id)
        self._refresh_models(None)
        self._reload_gen_tree()

    def _model_add(self):
        if not self.current_brand_id: return show_error("Спочатку виберіть бренд.")
        raw = self.model_entry.get()
        names = split_catalog_input(raw)
        if not names: return show_error("Введіть назву моделі (через кому для декількох).")
        for name in names:
            add_model(self.current_brand_id, name)
        self.model_entry.delete(0, tk.END)
        self._refresh_models(self.current_brand_id)
        self._reload_gen_tree()

    def _model_rename(self):
        sel = self.model_tree.selection()
        if not sel: return show_error("Виберіть модель.")
        model_id = int(sel[0].split("_")[1])
        name = self.model_entry.get().strip()
        if not name: return show_error("Введіть нову назву моделі.")
        result = rename_model(model_id, name)
        if result is not True:
            if isinstance(result, sqlite3.IntegrityError):
                show_error("Модель з такою назвою вже існує.")
            else:
                show_error("Не вдалося перейменувати модель.")
        self._refresh_models(self.current_brand_id)
        self.after(10, lambda: self._restore_tree_selection("model", f"model_{model_id}"))
        self._reload_gen_tree()

    def _model_delete(self):
        selection = list(self.model_tree.selection())
        if not selection:
            return show_error("Виберіть модель.")
        prompt = (
            "Видалити вибрану модель?"
            if len(selection) == 1
            else "Видалити вибрані моделі?"
        )
        if not messagebox.askyesno("Підтвердження", prompt):
            return
        ids = [int(iid.split("_")[1]) for iid in selection]
        brand_id = self.current_brand_id
        for model_id in ids:
            delete_model(model_id)
        self._refresh_models(brand_id)
        self._reload_gen_tree()

    def _open_specs(self):
        sel = self.model_tree.selection()
        if not sel: return show_error("Виберіть модель.")
        model_id = int(sel[0].split("_")[1])
        model_name = self.model_tree.item(sel[0], "values")[0]
        SpecsWindow(self, model_id, model_name)

    # -------- Шаблони
    def _build_tab_templates(self):
        wrap = ctk.CTkFrame(self.tab_templates)
        wrap.pack(fill="both", expand=True, padx=10, pady=10)

        selector = ctk.CTkFrame(wrap)
        selector.pack(fill="x", padx=10, pady=(6, 2))
        ctk.CTkLabel(selector, text="Категорія:").pack(side="left", padx=(0, 6))
        self.template_category_var = tk.StringVar(value="")
        self.template_category_menu = ctk.CTkOptionMenu(
            selector,
            values=["—"],
            variable=self.template_category_var,
            width=220,
            command=lambda _value: self._on_template_scope_change(),
        )
        self.template_category_menu.pack(side="left", padx=(0, 10))

        ctk.CTkLabel(selector, text="Тип плівки:").pack(side="left", padx=(0, 6))
        self.template_film_var = tk.StringVar(value="")
        self.template_film_menu = ctk.CTkOptionMenu(
            selector,
            values=["—"],
            variable=self.template_film_var,
            width=220,
            command=lambda _value: self._on_template_scope_change(),
        )
        self.template_film_menu.pack(side="left")

        ctk.CTkLabel(selector, text="Мова:").pack(side="left", padx=(10, 6))
        self.template_language_var = tk.StringVar(value="")
        self.template_language_menu = ctk.CTkOptionMenu(
            selector,
            values=["—"],
            variable=self.template_language_var,
            width=180,
            command=lambda _value: self._on_template_scope_change(),
        )
        self.template_language_menu.pack(side="left")

        # Ліва колонка: Заголовок і Теги
        left = ctk.CTkFrame(wrap)
        left.pack(side="left", fill="both", expand=True, padx=(0, 10), pady=5)

        ctk.CTkLabel(left, text="Шаблон заголовка ({{ brand }}, {{ model }}, {{ film_type }})").pack(anchor="w", padx=10, pady=(10, 0))
        self.title_box = ctk.CTkTextbox(left, height=80)
        self.title_box.pack(fill="x", padx=10, pady=5)
        self._bind_clipboard_shortcuts(self.title_box)
        self._bind_clipboard_context_menu(self.title_box)

        ctk.CTkLabel(left, text="Шаблон тегів ({{ brand }}, {{ model }}, {{ film_type }})").pack(anchor="w", padx=10, pady=(10, 0))
        self.tags_box = ctk.CTkTextbox(left, height=110)
        self.tags_box.pack(fill="x", padx=10, pady=5)
        self._bind_clipboard_shortcuts(self.tags_box)
        self._bind_clipboard_context_menu(self.tags_box)

        ctk.CTkButton(left, text="Зберегти заголовок/теги", command=self._save_title_tags).pack(anchor="e", padx=10, pady=10)

        # Права колонка: Опис для Категорії + Типу плівки
        right = ctk.CTkFrame(wrap)
        right.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=5)

        ctk.CTkLabel(right, text="Шаблон опису (доступні {{ brand }}, {{ model }}, {{ film_type }})").pack(anchor="w", padx=10, pady=(10, 0))
        editor_toolbar = ctk.CTkFrame(right)
        editor_toolbar.pack(fill="x", padx=10, pady=(6, 0))
        can_use_web_editor = self._desc_editor_ready
        self.desc_editor_btn = ctk.CTkButton(
            editor_toolbar,
            text="Відкрити візуальний редактор",
            command=self._open_desc_editor,
            state="normal" if can_use_web_editor else "disabled",
        )
        self.desc_editor_btn.pack(side="left")
        status_frame = ctk.CTkFrame(editor_toolbar, fg_color="transparent")
        status_frame.pack(side="left", fill="x", expand=True, padx=12)
        self.desc_editor_status_label = ctk.CTkLabel(
            status_frame,
            text="Перевіряємо стан веб-редактора…",
            justify="left",
            wraplength=360,
        )
        self.desc_editor_status_label.pack(side="left", fill="x", expand=True)
        self.desc_editor_retry_btn = ctk.CTkButton(
            status_frame,
            text="Спробувати ще раз",
            width=160,
            command=lambda: self._start_desc_editor_prepare(force=True),
        )
        self.desc_editor_retry_btn.pack(side="right", padx=(12, 0))
        self.desc_editor_retry_btn.pack_forget()
        if can_use_web_editor:
            self._set_desc_editor_status(
                "Редактор готовий до запуску у браузері. Після збереження поверніться до застосунку."
            )
        else:
            self._set_desc_editor_status(
                "Перевіряємо стан веб-редактора… Зачекайте, ми зберемо все автоматично."
            )

        self.desc_box = ctk.CTkTextbox(right, fg_color="#ffffff", text_color="#1f2933")
        self.desc_box.pack(fill="both", expand=True, padx=10, pady=5)
        self._bind_clipboard_shortcuts(self.desc_box)
        self._bind_clipboard_context_menu(self.desc_box)

        btn_row = ctk.CTkFrame(right)
        btn_row.pack(fill="x", padx=10, pady=6)
        self.desc_save_button = ctk.CTkButton(btn_row, text="Зберегти опис", command=self._save_desc_template)
        self.desc_save_button.pack(side="right")

        self.desc_cat_var = tk.StringVar(value="")

        self._refresh_template_selectors()
        self._start_desc_editor_prepare()

    def _save_title_tags(self, show_message: bool = True):
        film = self._selected_film_type_key()
        category_key = self._current_template_category
        title_value = self.title_box.get("1.0", "end").strip()
        tags_value = self.tags_box.get("1.0", "end").strip()

        language_code = self._current_template_language

        self._set_title_tags_block(category_key, film, language_code, title_value, tags_value)
        save_title_tags_templates(self.title_tags_templates)

        if category_key is None and film == "default" and not language_code:
            self.templates["title_template"] = title_value
            self.templates["tags_template"] = tags_value
            save_templates(self.templates)

        if show_message:
            show_info("Шаблони заголовку та тегів збережено.")

    def _load_title_tags_template(self):
        if not hasattr(self, "title_box") or not hasattr(self, "tags_box"):
            return
        category = self._current_template_category
        film = self._selected_film_type_key()
        language_codes = self._template_language_codes()
        language_code = self._current_template_language
        title_map, tags_map = resolve_title_tags(
            self.title_tags_templates,
            self.templates,
            category,
            film,
            language_codes,
        )

        def _pick(values: dict) -> str:
            if isinstance(language_code, str) and language_code in values:
                return values.get(language_code) or ""
            if None in values:
                return values.get(None) or ""
            for value in values.values():
                if isinstance(value, str):
                    return value
            return ""

        title_template = _pick(title_map)
        tags_template = _pick(tags_map)
        self.title_box.delete("1.0", "end")
        self.title_box.insert("1.0", title_template)
        self.tags_box.delete("1.0", "end")
        self.tags_box.insert("1.0", tags_template)

    def _on_template_scope_change(self, _selected_label=None):
        if not hasattr(self, "template_category_var") or not hasattr(self, "template_film_var"):
            return
        category_label = self.template_category_var.get()
        film_label = self.template_film_var.get()
        category_key = self._template_category_label_to_key.get(category_label)
        film_key = self._template_film_label_to_key.get(film_label, "default")
        if not film_key:
            film_key = "default"
        language_label = None
        language_code = None
        if hasattr(self, "template_language_var"):
            language_label = self.template_language_var.get()
            language_code = self._template_language_label_to_code.get(language_label)
        self._current_template_category = category_key
        self._current_film_type_key = film_key
        self._current_template_language = language_code
        target_category = category_key if category_key else GLOBAL_DESCRIPTION_KEY
        self._current_desc_category = target_category
        if hasattr(self, "desc_cat_var"):
            self.desc_cat_var.set(target_category)
        self._load_title_tags_template()
        self._load_desc_template()

    def _resolve_desc_template_html(self, category: Optional[str], film: str, language_code: Optional[str]):
        category_key = category or GLOBAL_DESCRIPTION_KEY
        film_key = film if film and film != "default" else "default"
        changed = False

        descs_by_category = self.templates.get("descriptions", {})
        if not isinstance(descs_by_category, dict):
            descs_by_category = {}
            self.templates["descriptions"] = descs_by_category
            changed = True

        descs = descs_by_category.get(category_key)
        if not isinstance(descs, dict):
            descs = {}
            descs_by_category[category_key] = descs
            changed = True

        if category_key != GLOBAL_DESCRIPTION_KEY:
            global_descs = descs_by_category.get(GLOBAL_DESCRIPTION_KEY)
            if not isinstance(global_descs, dict):
                global_descs = {}
                descs_by_category[GLOBAL_DESCRIPTION_KEY] = global_descs
                changed = True
        else:
            global_descs = descs

        sources = [descs]
        if category_key != GLOBAL_DESCRIPTION_KEY:
            sources.append(global_descs)

        raw_entry = None
        raw_store = None
        for store in sources:
            if isinstance(store, dict):
                candidate = store.get(film_key)
                if candidate is not None:
                    raw_entry = candidate
                    raw_store = store
                    break

        fallback_entry = None
        fallback_store = None
        if film_key != "default":
            for store in sources:
                if isinstance(store, dict):
                    candidate = store.get("default")
                    if candidate is not None:
                        fallback_entry = candidate
                        fallback_store = store
                        break

        normalized_language = language_code if language_code else None

        def _resolve_entry(entry, key=None, store=None):
            nonlocal changed
            value = None
            if isinstance(entry, dict):
                normalized = _normalize_template_language_entry(entry)
                if (
                    key is not None
                    and store is not None
                    and isinstance(store, dict)
                    and normalized is not entry
                ):
                    store[key] = normalized
                    changed = True
                value = _get_language_template_value(normalized, normalized_language, fallback_value=None)
                if value is None and normalized_language:
                    value = _get_language_template_value(normalized, None, fallback_value=None)
            elif isinstance(entry, str):
                value = entry
            return value

        html = _resolve_entry(raw_entry, key=film_key, store=raw_store)
        if html is None and fallback_entry is not None:
            html = _resolve_entry(fallback_entry, key="default", store=fallback_store)
        if html is None:
            html = ""
        return html, changed

    def _load_desc_template(self):
        if not hasattr(self, "desc_box"):
            return
        category = getattr(self, "_current_desc_category", None) or GLOBAL_DESCRIPTION_KEY
        film = self._selected_film_type_key()
        if hasattr(self, "desc_save_button"):
            self.desc_save_button.configure(state="normal")
        if hasattr(self, "desc_cat_var"):
            self.desc_cat_var.set(category)
        self._current_desc_category = category
        html, changed = self._resolve_desc_template_html(category, film, self._current_template_language)
        if changed:
            save_templates(self.templates)
        self.desc_box.configure(state="normal")
        self.desc_box.delete("1.0", "end")
        self.desc_box.insert("1.0", html)
        self._last_desc_html = html
        if hasattr(self, "desc_editor_btn"):
            can_use_web_editor = DESC_EDITOR_ENTRY.exists()
            self.desc_editor_btn.configure(state="normal" if can_use_web_editor else "disabled")

    def _apply_desc_editor_result(self, category: str, film: str, docs: Dict[str, Dict[str, object]]):
        descs_by_category = self.templates.setdefault("descriptions", {})
        if not isinstance(descs_by_category, dict):
            descs_by_category = {}
            self.templates["descriptions"] = descs_by_category
        film_map = descs_by_category.setdefault(category, {})
        if not isinstance(film_map, dict):
            film_map = {}
            descs_by_category[category] = film_map
        entry = film_map.get(film)
        changed = False
        for lang, doc in docs.items():
            if not isinstance(doc, dict):
                continue
            html_value = doc.get("html")
            if html_value is None:
                continue
            html_text = str(html_value).strip()
            lang_code = lang if isinstance(lang, str) and lang else None
            entry = _set_language_template_value(entry, lang_code, html_text, fallback_value="")
            changed = True
        if changed:
            film_map[film] = entry
            save_templates(self.templates)
            self._load_desc_template()
            show_info("Шаблон опису збережено.")

    def _on_desc_editor_finished(self) -> None:
        self._active_desc_host = None
        if hasattr(self, "desc_editor_btn"):
            can_use_web_editor = DESC_EDITOR_ENTRY.exists()
            self.desc_editor_btn.configure(state="normal" if can_use_web_editor else "disabled")

    def _set_desc_editor_status(self, message: str) -> None:
        if hasattr(self, "desc_editor_status_label"):
            try:
                self.desc_editor_status_label.configure(text=message)
            except Exception:  # pragma: no cover - UI best effort
                logger.exception("Не вдалося оновити статус веб-редактора")

    def _toggle_desc_editor_retry(self, visible: bool) -> None:
        if not hasattr(self, "desc_editor_retry_btn"):
            return
        if visible and not self._desc_editor_retry_visible:
            try:
                self.desc_editor_retry_btn.pack(side="right", padx=(12, 0))
                self._desc_editor_retry_visible = True
            except Exception:  # pragma: no cover - UI best effort
                logger.exception("Не вдалося показати кнопку повторної спроби збірки")
        elif not visible and self._desc_editor_retry_visible:
            try:
                self.desc_editor_retry_btn.pack_forget()
                self._desc_editor_retry_visible = False
            except Exception:
                logger.exception("Не вдалося приховати кнопку повторної спроби збірки")

    def _start_desc_editor_prepare(self, force: bool = False) -> None:
        thread = getattr(self, "_desc_editor_prepare_thread", None)
        if thread and thread.is_alive():
            return
        if force:
            self._desc_editor_error_shown = False
        self._set_desc_editor_status("Готуємо веб-редактор… Це може зайняти кілька хвилин.")
        if hasattr(self, "desc_editor_btn"):
            self.desc_editor_btn.configure(state="disabled")
        self._toggle_desc_editor_retry(False)

        def worker() -> None:
            try:
                ensure_desc_editor_built(force=force, quiet=True)
            except DescEditorBuildError as exc:
                logger.exception("Не вдалося автоматично зібрати веб-редактор", exc_info=True)
                self.after(0, lambda: self._on_desc_editor_prepare_failed(str(exc)))
                return
            self.after(0, self._on_desc_editor_prepare_ready)

        thread = threading.Thread(target=worker, daemon=True)
        self._desc_editor_prepare_thread = thread
        thread.start()

    def _on_desc_editor_prepare_ready(self) -> None:
        self._desc_editor_prepare_thread = None
        self._desc_editor_ready = True
        if hasattr(self, "desc_editor_btn"):
            self.desc_editor_btn.configure(state="normal")
        self._toggle_desc_editor_retry(False)
        self._set_desc_editor_status(
            "Редактор готовий до запуску у браузері. Після збереження поверніться до застосунку."
        )

    def _on_desc_editor_prepare_failed(self, error: str) -> None:
        self._desc_editor_prepare_thread = None
        self._desc_editor_ready = DESC_EDITOR_ENTRY.exists()
        if hasattr(self, "desc_editor_btn"):
            state = "normal" if self._desc_editor_ready else "disabled"
            self.desc_editor_btn.configure(state=state)
        self._toggle_desc_editor_retry(True)
        self._set_desc_editor_status(
            "Автозбірка не вдалася. Запустіть у терміналі 'python -m desc_editor_build --force' та повторіть спробу."
        )
        if not self._desc_editor_error_shown:
            self._desc_editor_error_shown = True
            show_error(
                "Автозбірка веб-редактора не вдалася.\n"
                "Виконайте 'python -m desc_editor_build --force' і переконайтеся, що встановлено Node.js 18+.\n"
                f"Деталі: {error}"
            )

    def _poll_desc_editor_result(
        self,
        host: DescriptionEditorHost,
        category: str,
        film: str,
        delay_ms: int = 600,
    ) -> None:
        result = host.poll_result()
        if result is None:
            if host.is_running:
                self.after(delay_ms, lambda: self._poll_desc_editor_result(host, category, film, delay_ms))
                return
            host.close()
            self._on_desc_editor_finished()
            return

        host.close()
        self._on_desc_editor_finished()
        if isinstance(result, dict) and result:
            self._apply_desc_editor_result(category, film, result)

    def _open_desc_editor(self):
        thread = getattr(self, "_desc_editor_prepare_thread", None)
        if thread and thread.is_alive():
            show_info("Веб-редактор ще готується. Зачекайте завершення автозбірки.")
            return
        if not getattr(self, "_desc_editor_ready", False):
            self._start_desc_editor_prepare(force=True)
            show_info("Редактор ще не готовий. Спробуйте ще раз за кілька хвилин.")
            return
        category = getattr(self, "_current_desc_category", None) or GLOBAL_DESCRIPTION_KEY
        film = self._selected_film_type_key()
        language_codes = list(dict.fromkeys([code for code in ["uk", "ru", "en"] + self._template_language_codes() if code]))
        if not language_codes:
            language_codes = ["uk"]
        existing_host = getattr(self, "_active_desc_host", None)
        if isinstance(existing_host, DescriptionEditorHost):
            try:
                existing_host.close()
            except Exception:
                logger.exception("Не вдалося закрити попередню сесію редактора опису")
            self._on_desc_editor_finished()
        docs: Dict[str, Dict[str, object]] = {}
        for lang in language_codes:
            html, _ = self._resolve_desc_template_html(category, film, lang)
            docs[lang] = {
                "lang": lang,
                "html": html,
                "css": "",
                "assets": [],
            }
        active_lang = self._current_template_language or language_codes[0]
        host = DescriptionEditorHost(docs, active_lang)
        try:
            host.launch()
        except DescriptionEditorError as exc:
            show_error(str(exc))
            return

        self._active_desc_host = host
        show_info(
            "Редактор відкрито у браузері. Після завершення натисніть 'Зберегти в застосунок' у вкладці браузера."
        )
        self._poll_desc_editor_result(host, category, film)

    def _save_desc_template(self):
        category = getattr(self, "_current_desc_category", None)
        if not category or category == GLOBAL_DESCRIPTION_KEY:
            template_category = getattr(self, "_current_template_category", None)
            if template_category:
                category = template_category
        if not category:
            category = self.desc_cat_var.get()
        if not category:
            category = GLOBAL_DESCRIPTION_KEY
        film = self._selected_film_type_key()
        txt = self.desc_box.get("1.0", "end").strip()
        language_code = self._current_template_language
        descs_by_category = self.templates.setdefault("descriptions", {})
        if not isinstance(descs_by_category, dict):
            descs_by_category = {}
            self.templates["descriptions"] = descs_by_category
        film_map = descs_by_category.setdefault(category, {})
        if not isinstance(film_map, dict):
            film_map = {}
            descs_by_category[category] = film_map
        entry = film_map.get(film)
        entry = _set_language_template_value(entry, language_code, txt, fallback_value="")
        film_map[film] = entry
        save_templates(self.templates)
        show_info("Шаблон опису збережено.")

    # -------- Параметри (мови + типи плівок)
    def _build_tab_parameters(self):
        wrap = ctk.CTkFrame(self.tab_parameters)
        wrap.pack(fill="both", expand=True, padx=10, pady=10)
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_columnconfigure(1, weight=1)
        wrap.grid_rowconfigure(0, weight=1)

        # ---- Мови шаблонів
        lang_column = ctk.CTkFrame(wrap)
        lang_column.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        lang_column.grid_columnconfigure(0, weight=1)
        lang_column.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(lang_column, text="Мови шаблонів").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 6))

        lang_tree_wrap = ctk.CTkFrame(lang_column)
        lang_tree_wrap.grid(row=1, column=0, sticky="nsew", padx=10)
        lang_tree_wrap.grid_columnconfigure(0, weight=1)
        lang_tree_wrap.grid_rowconfigure(0, weight=1)

        self.language_tree = ttk.Treeview(
            lang_tree_wrap,
            columns=("code", "label"),
            show="headings",
            selectmode="browse",
            height=8,
        )
        self.language_tree.heading("code", text="Код")
        self.language_tree.heading("label", text="Назва")
        self.language_tree.column("code", width=90, anchor="w")
        self.language_tree.column("label", width=200, anchor="w")
        self.language_tree.grid(row=0, column=0, sticky="nsew")

        lang_scroll = ttk.Scrollbar(lang_tree_wrap, orient="vertical", command=self.language_tree.yview)
        lang_scroll.grid(row=0, column=1, sticky="ns")
        self.language_tree.configure(yscrollcommand=lang_scroll.set)
        self.language_tree.bind("<<TreeviewSelect>>", self._on_language_select)

        lang_btns = ctk.CTkFrame(lang_column)
        lang_btns.grid(row=2, column=0, sticky="ew", padx=10, pady=(6, 0))
        ctk.CTkButton(lang_btns, text="Додати", command=self._language_add).pack(side="left", padx=4)
        ctk.CTkButton(lang_btns, text="Видалити", command=self._language_delete).pack(side="left", padx=4)

        lang_detail = ctk.CTkFrame(lang_column)
        lang_detail.grid(row=3, column=0, sticky="ew", padx=10, pady=(8, 10))
        lang_detail.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(lang_detail, text="Код мови").grid(row=0, column=0, sticky="w", pady=(0, 4))
        self.language_code_var = tk.StringVar(value="")
        self.language_code_entry = ctk.CTkEntry(lang_detail, textvariable=self.language_code_var)
        self.language_code_entry.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        self._bind_clipboard_shortcuts(self.language_code_entry)

        ctk.CTkLabel(lang_detail, text="Назва мови").grid(row=2, column=0, sticky="w", pady=(0, 4))
        self.language_label_var = tk.StringVar(value="")
        self.language_label_entry = ctk.CTkEntry(lang_detail, textvariable=self.language_label_var)
        self.language_label_entry.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        self._bind_clipboard_shortcuts(self.language_label_entry)

        ctk.CTkButton(lang_detail, text="Застосувати", command=self._language_apply).grid(row=4, column=0, sticky="e")

        # ---- Типи плівок
        film_wrap = ctk.CTkFrame(wrap)
        film_wrap.grid(row=0, column=1, sticky="nsew")
        film_wrap.grid_columnconfigure(1, weight=1)
        film_wrap.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(film_wrap, text="Типи плівок").grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 6))

        list_frame = ctk.CTkFrame(film_wrap)
        list_frame.grid(row=1, column=0, sticky="ns", padx=(0, 10))
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        self.filmtype_tree = ttk.Treeview(
            list_frame,
            columns=("name", "enabled"),
            show="headings",
            selectmode="browse",
            height=14,
        )
        self.filmtype_tree.heading("name", text="Назва")
        self.filmtype_tree.heading("enabled", text="Увімкнено")
        self.filmtype_tree.column("name", width=220, anchor="w")
        self.filmtype_tree.column("enabled", width=90, anchor="center")
        self.filmtype_tree.grid(row=0, column=0, sticky="nsew")

        y_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.filmtype_tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        self.filmtype_tree.configure(yscrollcommand=y_scroll.set)
        self.filmtype_tree.bind("<<TreeviewSelect>>", self._on_filmtype_select)

        btn_frame = ctk.CTkFrame(list_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ctk.CTkButton(btn_frame, text="Додати", command=self._filmtype_add).pack(side="left", padx=4)
        ctk.CTkButton(btn_frame, text="Видалити", command=self._filmtype_delete).pack(side="left", padx=4)

        detail = ctk.CTkFrame(film_wrap)
        detail.grid(row=1, column=1, sticky="nsew")
        detail.grid_rowconfigure(3, weight=1)

        ctk.CTkLabel(detail, text="Назва").pack(anchor="w", padx=10, pady=(10, 0))
        self.filmtype_name_entry = ctk.CTkEntry(detail, textvariable=self.filmtype_name_var)
        self.filmtype_name_entry.pack(fill="x", padx=10, pady=(0, 8))
        self._bind_clipboard_shortcuts(self.filmtype_name_entry)

        self.filmtype_enabled_check = ctk.CTkCheckBox(
            detail,
            text="Увімкнено за замовчуванням",
            variable=self.filmtype_enabled_var,
        )
        self.filmtype_enabled_check.pack(anchor="w", padx=10, pady=(0, 8))

        ctk.CTkButton(detail, text="Застосувати", command=self._filmtype_apply).pack(anchor="e", padx=10, pady=10)

        self._refresh_language_tree(select_index=0 if self.templates.get("template_languages") else None)
        self._refresh_filmtype_tree(select_index=0 if self.templates.get("film_types") else None)

    def _refresh_language_tree(self, select_index=None):
        tree = getattr(self, "language_tree", None)
        if tree is None:
            return
        raw_languages = self.templates.get("template_languages")
        if not isinstance(raw_languages, list):
            raw_languages = []
        normalized = _normalize_language_definitions(raw_languages)
        if normalized != raw_languages:
            self.templates["template_languages"] = normalized
            save_templates(self.templates)
        tree.delete(*tree.get_children())
        for idx, item in enumerate(normalized):
            code = (item.get("code") or "").strip()
            label = (item.get("label") or "").strip()
            tree.insert("", "end", iid=f"lang_{idx}", values=(code, label))
        if select_index is not None and 0 <= select_index < len(normalized):
            iid = f"lang_{select_index}"
            if tree.exists(iid):
                tree.selection_set(iid)
                tree.focus(iid)
                tree.see(iid)
                self._on_language_select()
        elif not normalized:
            self._current_language_index = None
            if hasattr(self, "language_code_var"):
                self.language_code_var.set("")
            if hasattr(self, "language_label_var"):
                self.language_label_var.set("")

    def _on_language_select(self, _evt=None):
        tree = getattr(self, "language_tree", None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            self._current_language_index = None
            if hasattr(self, "language_code_var"):
                self.language_code_var.set("")
            if hasattr(self, "language_label_var"):
                self.language_label_var.set("")
            return
        iid = sel[0]
        parts = iid.split("_")
        if len(parts) != 2:
            return
        try:
            idx = int(parts[1])
        except ValueError:
            return
        languages = self.templates.get("template_languages", [])
        if idx < 0 or idx >= len(languages):
            return
        self._current_language_index = idx
        entry = languages[idx]
        code = (entry.get("code") or "").strip()
        label = (entry.get("label") or "").strip()
        if hasattr(self, "language_code_var"):
            self.language_code_var.set(code)
        if hasattr(self, "language_label_var"):
            self.language_label_var.set(label)

    def _language_add(self):
        code = simpledialog.askstring("Нова мова", "Введіть код мови (наприклад, uk):", parent=self)
        if code is None:
            return
        code = code.strip()
        if not code:
            return show_error("Введіть код мови.")
        languages = self.templates.setdefault("template_languages", [])
        existing_codes = {str(item.get("code", "")).strip().lower() for item in languages}
        if code.lower() in existing_codes:
            return show_error("Мова з таким кодом вже існує.")
        label = simpledialog.askstring("Нова мова", "Введіть назву мови:", initialvalue=code, parent=self)
        if label is None:
            return
        label = label.strip() or code
        languages.append({"code": code, "label": label})
        save_templates(self.templates)
        self._current_template_language = code
        self._on_languages_changed()
        self._refresh_language_tree(select_index=len(languages) - 1)

    def _language_delete(self):
        tree = getattr(self, "language_tree", None)
        if tree is None:
            return
        selection = list(tree.selection())
        if not selection:
            return show_error("Виберіть мову для видалення.")
        if not messagebox.askyesno("Підтвердження", "Видалити вибрану мову?"):
            return
        indices = []
        for iid in selection:
            parts = iid.split("_")
            if len(parts) != 2:
                continue
            try:
                idx = int(parts[1])
            except ValueError:
                continue
            indices.append(idx)
        indices = sorted(set(indices), reverse=True)
        removed_codes = []
        languages = self.templates.get("template_languages", [])
        removed_current = False
        for idx in indices:
            if 0 <= idx < len(languages):
                entry = languages.pop(idx)
                code = (entry.get("code") or "").strip()
                if code:
                    removed_codes.append(code)
                    if code == self._current_template_language:
                        removed_current = True
        if removed_codes:
            save_templates(self.templates)
            for code in removed_codes:
                self._update_language_code_references(code, None)
            if removed_current:
                self._current_template_language = None
            self._on_languages_changed()
        self._refresh_language_tree(select_index=None)

    def _language_apply(self):
        idx = self._current_language_index
        languages = self.templates.get("template_languages", [])
        if idx is None or idx < 0 or idx >= len(languages):
            return show_error("Виберіть мову для редагування.")
        code = self.language_code_var.get().strip()
        label = self.language_label_var.get().strip()
        if not code:
            return show_error("Код мови не може бути порожнім.")
        if not label:
            label = code
        old_entry = languages[idx]
        old_code = (old_entry.get("code") or "").strip()
        if code.lower() != old_code.lower():
            existing_codes = {str(item.get("code", "")).strip().lower() for i, item in enumerate(languages) if i != idx}
            if code.lower() in existing_codes:
                return show_error("Мова з таким кодом вже існує.")
        languages[idx] = {"code": code, "label": label}
        save_templates(self.templates)
        if code != old_code:
            self._update_language_code_references(old_code, code)
        if old_code == self._current_template_language:
            self._current_template_language = code
        self._on_languages_changed()
        self._refresh_language_tree(select_index=idx)

    def _update_language_code_references(self, old_code: str, new_code: Optional[str]):
        if not isinstance(old_code, str) or not old_code:
            return
        changed_title_tags = False

        def adjust_title_block(block):
            nonlocal changed_title_tags
            if not isinstance(block, dict):
                return
            for key in ("title_template", "tags_template"):
                entry = block.get(key)
                if _rename_language_in_entry(entry, old_code, new_code):
                    changed_title_tags = True

        adjust_title_block(self.title_tags_templates.get("default"))
        by_film = self.title_tags_templates.get("by_film")
        if isinstance(by_film, dict):
            for film_block in by_film.values():
                adjust_title_block(film_block)
        by_category = self.title_tags_templates.get("by_category")
        if isinstance(by_category, dict):
            for cat_entry in by_category.values():
                if not isinstance(cat_entry, dict):
                    continue
                adjust_title_block(cat_entry.get("default"))
                films = cat_entry.get("by_film")
                if isinstance(films, dict):
                    for film_block in films.values():
                        adjust_title_block(film_block)

        changed_descriptions = False
        descriptions = self.templates.get("descriptions")
        if isinstance(descriptions, dict):
            for cat_key, film_map in descriptions.items():
                if not isinstance(film_map, dict):
                    continue
                for film_key, entry in list(film_map.items()):
                    if isinstance(entry, dict):
                        normalized = _normalize_template_language_entry(entry)
                        if normalized is not entry:
                            film_map[film_key] = normalized
                            entry = normalized
                            changed_descriptions = True
                        if _rename_language_in_entry(entry, old_code, new_code):
                            changed_descriptions = True

        changed_export_fields = False
        for field in self.export_fields:
            if not isinstance(field, dict):
                continue
            languages_value = field.get("languages")
            if isinstance(languages_value, str):
                if languages_value.strip() == old_code:
                    if new_code:
                        field["languages"] = [new_code]
                    else:
                        field["languages"] = []
                    changed_export_fields = True
            elif isinstance(languages_value, (list, tuple, set)):
                updated_list = []
                modified = False
                for lang in languages_value:
                    if not isinstance(lang, str):
                        continue
                    stripped = lang.strip()
                    if stripped == old_code:
                        if new_code:
                            updated_list.append(new_code)
                        modified = True
                    else:
                        updated_list.append(stripped)
                if modified or len(updated_list) != len(languages_value):
                    field["languages"] = updated_list
                    changed_export_fields = True

        if changed_title_tags:
            save_title_tags_templates(self.title_tags_templates)
        if changed_descriptions:
            save_templates(self.templates)
        if changed_export_fields:
            save_export_fields(self.export_fields)

    def _on_languages_changed(self):
        self._refresh_template_selectors()
        self._refresh_export_language_controls()
        self._refresh_export_fields_tree()
        self._refresh_language_tree()

    def _refresh_export_language_controls(self):
        self._build_export_field_language_checkboxes()
        self._build_generate_language_checkboxes()
        selected_index = getattr(self, "_export_selected_index", None)
        if selected_index is not None:
            self._load_export_field_detail(selected_index)

    def _build_export_field_language_checkboxes(self):
        frame = getattr(self, "export_field_language_checks_frame", None)
        if frame is None:
            return
        for child in list(frame.winfo_children()):
            try:
                child.destroy()
            except Exception:
                pass
        self.export_field_language_vars = {}
        self.export_field_language_checks = []
        language_items = self._template_language_items()
        for label, code in language_items:
            if not code:
                continue
            var = tk.BooleanVar(value=False)
            checkbox = ctk.CTkCheckBox(frame, text=label, variable=var)
            checkbox.pack(side="left", padx=4, pady=2)
            self.export_field_language_vars[code] = var
            self.export_field_language_checks.append(checkbox)

    def _build_generate_language_checkboxes(self):
        container = getattr(self, "generate_language_checks_container", None)
        if container is None:
            return
        for child in list(container.winfo_children()):
            try:
                child.destroy()
            except Exception:
                pass
        self.export_language_vars = []
        language_items = self._template_language_items()
        codes = [code for label, code in language_items if code]
        if not codes:
            hint_label = getattr(self, "generate_language_hint", None)
            if hint_label is not None:
                hint_label.configure(text="")
            return
        for code in codes:
            label = self._language_label_for_code(code)
            var = tk.BooleanVar(value=True)
            checkbox = ctk.CTkCheckBox(container, text=label, variable=var)
            checkbox.pack(side="left", padx=6, pady=2)
            self.export_language_vars.append((code, var))
        hint_label = getattr(self, "generate_language_hint", None)
        if hint_label is not None:
            hint_label.configure(text="Залиште всі позначені, щоб експортувати всі мови.")


    def _refresh_filmtype_tree(self, select_index=None):
        tree = getattr(self, "filmtype_tree", None)
        if tree is None:
            return
        tree.delete(*tree.get_children())
        film_types = self.templates.get("film_types", [])
        for idx, item in enumerate(film_types):
            name = (item.get("name") or "").strip()
            enabled = "Так" if item.get("enabled", True) else "Ні"
            tree.insert("", "end", iid=f"ft_{idx}", values=(name, enabled))
        if select_index is not None and 0 <= select_index < len(film_types):
            iid = f"ft_{select_index}"
            if tree.exists(iid):
                tree.selection_set(iid)
                tree.focus(iid)
                tree.see(iid)
                self._on_filmtype_select()
        elif not film_types:
            self._current_filmtype_index = None
            self.filmtype_name_var.set("")
            self.filmtype_enabled_var.set(True)

    def _on_filmtype_select(self, _evt=None):
        tree = getattr(self, "filmtype_tree", None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            self._current_filmtype_index = None
            self.filmtype_name_var.set("")
            self.filmtype_enabled_var.set(True)
            return
        iid = sel[0]
        parts = iid.split("_")
        if len(parts) != 2:
            return
        try:
            idx = int(parts[1])
        except ValueError:
            return
        film_types = self.templates.get("film_types", [])
        if idx < 0 or idx >= len(film_types):
            return
        self._current_filmtype_index = idx
        item = film_types[idx]
        self.filmtype_name_var.set(item.get("name", ""))
        self.filmtype_enabled_var.set(bool(item.get("enabled", True)))

    def _filmtype_add(self):
        new_name = simpledialog.askstring("Новий тип плівки", "Введіть назву типу плівки:", parent=self)
        if new_name is None:
            return
        new_name = new_name.strip()
        if not new_name:
            return show_error("Введіть назву типу плівки.")
        existing = {((item.get("name") or "").strip().lower()) for item in self.templates.get("film_types", [])}
        if new_name.lower() in existing:
            return show_error("Тип плівки з такою назвою вже існує.")
        self.templates.setdefault("film_types", []).append({"name": new_name, "enabled": True})
        save_templates(self.templates)
        if self._ensure_title_tags_film(new_name):
            save_title_tags_templates(self.title_tags_templates)
        self._refresh_filmtype_tree(select_index=len(self.templates.get("film_types", [])) - 1)
        self._refresh_filmtype_checkboxes()
        self._refresh_template_selectors()

    def _filmtype_delete(self):
        tree = getattr(self, "filmtype_tree", None)
        if tree is None:
            return
        selection = list(tree.selection())
        if not selection:
            return show_error("Виберіть тип плівки.")
        if not messagebox.askyesno("Підтвердження", "Видалити вибраний тип плівки?"):
            return
        indices = []
        for iid in selection:
            parts = iid.split("_")
            if len(parts) != 2:
                continue
            try:
                idx = int(parts[1])
            except ValueError:
                continue
            indices.append(idx)
        indices = sorted(set(indices), reverse=True)
        removed_names = []
        for idx in indices:
            film_types = self.templates.get("film_types", [])
            if 0 <= idx < len(film_types):
                removed = film_types.pop(idx)
                name = (removed.get("name") or "").strip()
                if name:
                    removed_names.append(name)
        if removed_names:
            save_templates(self.templates)
            for name in removed_names:
                self._remove_film_type_templates(name)
                if self._current_film_type_key == name:
                    self._current_film_type_key = "default"
        self._current_filmtype_index = None
        self.filmtype_name_var.set("")
        self.filmtype_enabled_var.set(True)
        self._refresh_filmtype_tree(select_index=None)
        self._refresh_filmtype_checkboxes()
        self._refresh_template_selectors()

    def _filmtype_apply(self):
        idx = self._current_filmtype_index
        film_types = self.templates.get("film_types", [])
        if idx is None or idx < 0 or idx >= len(film_types):
            return show_error("Виберіть тип плівки.")
        new_name = self.filmtype_name_var.get().strip()
        if not new_name:
            return show_error("Введіть назву типу плівки.")
        for pos, item in enumerate(film_types):
            if pos != idx and (item.get("name") or "").strip().lower() == new_name.lower():
                return show_error("Тип плівки з такою назвою вже існує.")
        old_name = film_types[idx].get("name", "")
        film_types[idx]["name"] = new_name
        film_types[idx]["enabled"] = bool(self.filmtype_enabled_var.get())
        save_templates(self.templates)
        if new_name != old_name:
            self._rename_film_type(old_name, new_name)
            if self._current_film_type_key == old_name:
                self._current_film_type_key = new_name
        self._ensure_title_tags_film(new_name)
        save_title_tags_templates(self.title_tags_templates)
        self._refresh_filmtype_tree(select_index=idx)
        self._refresh_filmtype_checkboxes()
        self._refresh_template_selectors()
        show_info("Тип плівки оновлено.")

    # -------- Експортні поля
    def _build_tab_export(self):
        wrap = ctk.CTkFrame(self.tab_export)
        wrap.pack(fill="both", expand=True, padx=10, pady=10)
        wrap.grid_columnconfigure(0, weight=0)
        wrap.grid_columnconfigure(1, weight=1)
        wrap.grid_rowconfigure(0, weight=1)

        list_frame = ctk.CTkFrame(wrap)
        list_frame.grid(row=0, column=0, sticky="ns", padx=(0, 10))
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        self.export_fields_tree = ttk.Treeview(
            list_frame,
            columns=("field", "enabled"),
            show="headings",
            selectmode="browse",
            height=18,
        )
        self.export_fields_tree.heading("field", text="Поле")
        self.export_fields_tree.heading("enabled", text="Увімкнено")
        self.export_fields_tree.column("field", width=220, anchor="w")
        self.export_fields_tree.column("enabled", width=90, anchor="center")
        self.export_fields_tree.grid(row=0, column=0, sticky="nsew")

        y_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.export_fields_tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        self.export_fields_tree.configure(yscrollcommand=y_scroll.set)
        self.export_fields_tree.bind("<<TreeviewSelect>>", self._on_export_field_select)

        btn_frame = ctk.CTkFrame(list_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ctk.CTkButton(btn_frame, text="Додати поле", command=self._export_add_field).pack(side="left", padx=4)
        ctk.CTkButton(btn_frame, text="Видалити", command=self._export_delete_field).pack(side="left", padx=4)
        ctk.CTkButton(btn_frame, text="Вгору", command=lambda: self._export_move_field(-1)).pack(side="right", padx=4)
        ctk.CTkButton(btn_frame, text="Вниз", command=lambda: self._export_move_field(1)).pack(side="right", padx=4)

        detail = ctk.CTkFrame(wrap)
        detail.grid(row=0, column=1, sticky="nsew")
        detail.grid_rowconfigure(4, weight=1)

        self.export_field_name_var = tk.StringVar(value="")
        self.export_field_enabled_var = tk.BooleanVar(value=False)

        ctk.CTkLabel(detail, text="Назва поля").pack(anchor="w", padx=10, pady=(8, 0))
        self.export_field_name_entry = ctk.CTkEntry(detail, textvariable=self.export_field_name_var)
        self.export_field_name_entry.pack(fill="x", padx=10, pady=(0, 8))
        self._bind_clipboard_shortcuts(self.export_field_name_entry)

        self.export_field_enabled_check = ctk.CTkCheckBox(
            detail,
            text="Увімкнути поле",
            variable=self.export_field_enabled_var,
        )
        self.export_field_enabled_check.pack(anchor="w", padx=10, pady=(0, 8))

        lang_block = ctk.CTkFrame(detail, fg_color="transparent")
        lang_block.pack(fill="x", padx=10, pady=(0, 8))
        ctk.CTkLabel(lang_block, text="Мови поля:").pack(anchor="w", pady=(0, 4))
        lang_checks = ctk.CTkFrame(lang_block, fg_color="transparent")
        lang_checks.pack(fill="x")
        self.export_field_language_checks_frame = lang_checks
        self._build_export_field_language_checkboxes()
        self.export_language_hint_label = ctk.CTkLabel(
            lang_block,
            text="Без вибору мови поле застосовується до всіх.",
            anchor="w",
            justify="left",
            wraplength=360,
            font=ctk.CTkFont(size=12),
        )
        self.export_language_hint_label.pack(fill="x", pady=(4, 0))
        self._export_language_hint_default = self.export_language_hint_label.cget("text")

        ctk.CTkLabel(detail, text="Шаблон (формула =IF(...) або Jinja2)").pack(anchor="w", padx=10, pady=(0, 4))
        self.export_field_template = ctk.CTkTextbox(detail, height=220)
        self.export_field_template.pack(fill="both", expand=True, padx=10, pady=(0, 6))
        self.export_field_template.bind("<<Modified>>", self._on_export_template_modified)
        self._bind_clipboard_shortcuts(self.export_field_template)

        hint_text = (
            "Почніть рядок зі знаком '=' щоб використати формулу у стилі Google Sheets (аргументи через ';').\n"
            "Без '=' шаблон рендериться через Jinja2. Доступні змінні: {{ brand }}, {{ model }}, {{ category }}, {{ film_type }}, "
            "{{ title }}, {{ description }}, {{ tags }}, {{ spec('Назва') }}, {{ specs['Ключ'] }}, {{ row_number }}, {{ now }}."
        )
        ctk.CTkLabel(detail, text=hint_text, justify="left", anchor="w", wraplength=360).pack(
            fill="x", padx=10, pady=(0, 4)
        )

        self.export_template_status = ctk.CTkLabel(
            detail,
            text="",
            justify="left",
            anchor="w",
            wraplength=360,
            text_color="#888888",
        )
        self.export_template_status.pack(fill="x", padx=10, pady=(0, 8))
        self._update_template_status("")

        action_row = ctk.CTkFrame(detail)
        action_row.pack(fill="x", padx=10, pady=(0, 10))
        self.export_apply_button = ctk.CTkButton(action_row, text="Застосувати", command=lambda: self._export_apply_detail(False))
        self.export_apply_button.pack(side="right", padx=4)
        ctk.CTkButton(action_row, text="Зберегти", command=self._export_save_all).pack(side="right", padx=4)
        ctk.CTkButton(action_row, text="Відновити стандартні", command=self._export_reset_defaults).pack(side="left", padx=4)

        self._refresh_export_fields_tree(select_index=0 if self.export_fields else None)
        if self.export_fields:
            self._load_export_field_detail(0)
        else:
            self._load_export_field_detail(None)

    def _build_tab_files(self):
        wrap = ctk.CTkFrame(self.tab_files)
        wrap.pack(fill="both", expand=True, padx=10, pady=10)
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_rowconfigure(1, weight=0)

        info_text = (
            "Експорт створює резервну копію всіх даних (каталог, шаблони, параметри, налаштування експорту) у форматі Excel.\n"
            "Імпорт замінює поточні дані значеннями з вибраного файлу."
        )
        ctk.CTkLabel(wrap, text=info_text, justify="left", wraplength=720).grid(
            row=0, column=0, sticky="w", padx=12, pady=(12, 8)
        )

        self.files_status_var = tk.StringVar(
            value="Створіть резервну копію перед імпортом, щоб уникнути втрати даних."
        )
        ctk.CTkLabel(wrap, textvariable=self.files_status_var, justify="left", wraplength=720).grid(
            row=1, column=0, sticky="w", padx=12, pady=(0, 10)
        )

        buttons = ctk.CTkFrame(wrap)
        buttons.grid(row=2, column=0, sticky="w", padx=12, pady=(0, 12))
        ctk.CTkButton(buttons, text="Експорт у Excel", command=self._files_export_all, width=200).pack(
            side="left", padx=(0, 10)
        )
        ctk.CTkButton(buttons, text="Імпорт з Excel", command=self._files_import_all, width=200).pack(
            side="left"
        )

    def _files_set_status(self, message: str):
        status_var = getattr(self, "files_status_var", None)
        if isinstance(status_var, tk.StringVar):
            status_var.set(message)

    def _files_export_all(self):
        self._export_apply_detail(False)
        self._save_title_tags(show_message=False)

        default_name = f"prom_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path = filedialog.asksaveasfilename(
            title="Зберегти резервну копію",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel", "*.xlsx"), ("Усі файли", "*.*")],
        )
        if not file_path:
            return

        try:
            export_all_data_to_excel(
                file_path,
                self.templates,
                self.title_tags_templates,
                self.export_fields,
            )
        except DataTransferError as exc:
            self._files_set_status(f"Помилка експорту: {exc.message}")
            return show_error(f"Помилка експорту (код {exc.code}): {exc.message}")
        except Exception as exc:  # pragma: no cover - defensive
            logger.exception("Unexpected error during data export")
            self._files_set_status("Сталася непередбачена помилка під час експорту.")
            return show_error(f"Не вдалося зберегти дані: {exc}")

        self._files_set_status(f"Дані збережено у файл: {file_path}")
        show_info("Резервну копію успішно створено.")

    def _files_import_all(self):
        self._export_apply_detail(False)

        file_path = filedialog.askopenfilename(
            title="Виберіть файл резервної копії",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xltx *.xltm"), ("Усі файли", "*.*")],
        )
        if not file_path:
            return

        if not messagebox.askyesno(
            "Підтвердження",
            "Імпорт повністю замінить поточні дані. Продовжити?",
        ):
            return

        try:
            result = import_all_data_from_excel(file_path)
        except DataTransferError as exc:
            self._files_set_status(f"Помилка імпорту: {exc.message}")
            return show_error(f"Помилка імпорту (код {exc.code}): {exc.message}")
        except Exception as exc:  # pragma: no cover - defensive
            logger.exception("Unexpected error during data import")
            self._files_set_status("Сталася непередбачена помилка під час імпорту.")
            return show_error(f"Не вдалося імпортувати дані: {exc}")

        self.templates = result.get("templates", self.templates)
        self.title_tags_templates = result.get("title_tags_templates", self.title_tags_templates)
        self.export_fields = result.get("export_fields", self.export_fields)

        self._refresh_categories()
        self._refresh_language_tree(select_index=0 if self.templates.get("template_languages") else None)
        self._refresh_filmtype_tree(select_index=0 if self.templates.get("film_types") else None)
        self._refresh_filmtype_checkboxes()
        self._refresh_template_selectors()
        self._load_title_tags_template()
        self._load_desc_template()
        self._refresh_export_fields_tree(select_index=0 if self.export_fields else None)
        if self.export_fields:
            self._load_export_field_detail(0)
        else:
            self._load_export_field_detail(None)

        self._files_set_status(f"Дані імпортовано з файлу: {file_path}")
        show_info("Дані успішно імпортовано.")

    def _refresh_export_fields_tree(self, select_index=None):
        tree = getattr(self, "export_fields_tree", None)
        if tree is None:
            return
        self._export_tree_updating = True
        tree.delete(*tree.get_children())
        for idx, field in enumerate(self.export_fields):
            name = str(field.get("field", "")).strip()
            languages = field.get("languages", [])
            display_name = name
            codes = []
            if isinstance(languages, str):
                lang_code = languages.strip()
                if lang_code:
                    codes.append(lang_code)
            elif isinstance(languages, (list, tuple, set)):
                seen_langs = set()
                for lang in languages:
                    if not isinstance(lang, str):
                        continue
                    code = lang.strip()
                    if not code or code in seen_langs:
                        continue
                    codes.append(code)
                    seen_langs.add(code)
            if codes:
                labels = [self._language_label_for_code(code) for code in codes]
                display_name = f"{name} ({', '.join(labels)})"
            status = "Так" if field.get("enabled") else "Ні"
            tree.insert("", "end", iid=f"exp_{idx}", values=(display_name, status))
        if select_index is not None and 0 <= select_index < len(self.export_fields):
            iid = f"exp_{select_index}"
            tree.selection_set(iid)
            tree.focus(iid)
        self._export_tree_updating = False

    def _set_export_detail_state(self, enabled: bool):
        state = tk.NORMAL if enabled else tk.DISABLED
        if hasattr(self, "export_field_name_entry"):
            self.export_field_name_entry.configure(state=state)
        if hasattr(self, "export_field_enabled_check"):
            self.export_field_enabled_check.configure(state=state)
        if hasattr(self, "export_field_language_checks"):
            for checkbox in self.export_field_language_checks:
                checkbox.configure(state=state)
        if hasattr(self, "export_apply_button"):
            self.export_apply_button.configure(state=state)
        if hasattr(self, "export_field_template"):
            self.export_field_template.configure(state="normal" if enabled else "disabled")
        self._update_template_status()

    def _load_export_field_detail(self, index):
        if index is None or index < 0 or index >= len(self.export_fields):
            self._export_selected_index = None
            self._set_export_detail_state(False)
            if hasattr(self, "export_field_name_var"):
                self.export_field_name_var.set("")
            if hasattr(self, "export_field_enabled_var"):
                self.export_field_enabled_var.set(False)
            if hasattr(self, "export_field_language_vars"):
                for var in self.export_field_language_vars.values():
                    try:
                        var.set(False)
                    except Exception:
                        pass
            self._export_unknown_language_codes = []
            if hasattr(self, "export_language_hint_label"):
                default_hint = getattr(self, "_export_language_hint_default", None)
                if default_hint is not None:
                    self.export_language_hint_label.configure(text=default_hint)
            if hasattr(self, "export_field_template"):
                self.export_field_template.configure(state="normal")
                self.export_field_template.delete("1.0", "end")
                self.export_field_template.configure(state="disabled")
            self._update_template_status("")
            return

        self._export_selected_index = index
        field = self.export_fields[index]
        name = str(field.get("field", ""))
        template = field.get("template", "")
        if template is None:
            template = ""
        enabled = bool(field.get("enabled"))

        self._set_export_detail_state(True)
        self.export_field_name_var.set(name)
        self.export_field_enabled_var.set(enabled)
        language_vars = getattr(self, "export_field_language_vars", {})
        unknown_codes = []
        if isinstance(language_vars, dict):
            for var in language_vars.values():
                try:
                    var.set(False)
                except Exception:
                    pass
        languages_value = field.get("languages", [])
        parsed_languages = []
        if isinstance(languages_value, str):
            code = languages_value.strip()
            if code:
                parsed_languages.append(code)
        elif isinstance(languages_value, (list, tuple, set)):
            seen_codes = set()
            for lang in languages_value:
                if not isinstance(lang, str):
                    continue
                code = lang.strip()
                if not code or code in seen_codes:
                    continue
                parsed_languages.append(code)
                seen_codes.add(code)
        if isinstance(language_vars, dict):
            for code in parsed_languages:
                var = language_vars.get(code)
                if var is not None:
                    try:
                        var.set(True)
                    except Exception:
                        pass
                else:
                    unknown_codes.append(code)
        self._export_unknown_language_codes = unknown_codes
        if hasattr(self, "export_language_hint_label"):
            default_hint = getattr(self, "_export_language_hint_default", "")
            if unknown_codes:
                extras = ", ".join(sorted(unknown_codes))
                hint_text = f"{default_hint}\nНевідомі коди збережено: {extras}"
            else:
                hint_text = default_hint
            self.export_language_hint_label.configure(text=hint_text)
        self.export_field_template.configure(state="normal")
        self.export_field_template.delete("1.0", "end")
        self.export_field_template.insert("1.0", str(template))
        self.export_field_template.edit_modified(False)
        self._update_template_status(template)

    def _on_export_field_select(self, _event):
        if getattr(self, "_export_tree_updating", False):
            return
        selection = self.export_fields_tree.selection() if hasattr(self, "export_fields_tree") else []
        if not selection:
            self._load_export_field_detail(None)
            return
        iid = selection[0]
        try:
            idx = int(iid.split("_", 1)[1])
        except (IndexError, ValueError):
            idx = None
        self._export_apply_detail(False)
        if idx is None:
            self._load_export_field_detail(None)
        else:
            self._load_export_field_detail(idx)

    def _on_export_template_modified(self, _event):
        widget = getattr(self, "export_field_template", None)
        if widget is None:
            return
        try:
            modified = bool(widget.edit_modified())
        except Exception:
            modified = True
        if not modified:
            return
        try:
            widget.edit_modified(False)
        except Exception:
            pass
        text = widget.get("1.0", "end").rstrip()
        self._update_template_status(text)

    def _update_template_status(self, template_text: str | None = None):
        label = getattr(self, "export_template_status", None)
        if label is None:
            return
        if template_text is None:
            widget = getattr(self, "export_field_template", None)
            if widget is None:
                template_text = ""
            else:
                template_text = widget.get("1.0", "end").rstrip()
        message, color = self._analyze_template_text(template_text)
        label.configure(text=message, text_color=color)

    def _analyze_template_text(self, template_text: str):
        trimmed = template_text.strip()
        if not trimmed:
            return (
                "Введіть формулу (=IF(...)) або шаблон Jinja2. Аргументи формули розділяйте ';'.",
                "#888888",
            )
        if _looks_like_formula(trimmed):
            try:
                info = FormulaEngine.describe(trimmed)
            except FormulaError as exc:
                return (f"❌ Помилка формули: {exc}", "#c94a4a")
            variables = sorted(info.get("variables", []))
            if variables:
                vars_text = ", ".join(variables)
                hint = f"Змінні: {vars_text}"
            else:
                hint = "Змінні не використовуються."
            return (f"✅ Формула валідна. {hint}", "#4c9a2a")
        try:
            Template(trimmed)
        except TemplateError as exc:
            return (f"❌ Помилка шаблону Jinja2: {exc}", "#c94a4a")
        return (
            "ℹ️ Використовується шаблон Jinja2. Доступні змінні: {{ brand }}, {{ model }}, {{ category }}, {{ film_type }}, {{ title }}, {{ description }}, {{ tags }}, {{ row_number }}, {{ now }}.",
            "#888888",
        )

    def _export_apply_detail(self, save_to_file: bool):
        idx = getattr(self, "_export_selected_index", None)
        if idx is None or idx < 0 or idx >= len(self.export_fields):
            return False
        field = self.export_fields[idx]
        name = self.export_field_name_var.get().strip()
        if not name:
            name = field.get("field") or f"Поле_{idx + 1}"
            self.export_field_name_var.set(name)
        template = self.export_field_template.get("1.0", "end").rstrip()
        enabled = bool(self.export_field_enabled_var.get())
        language_vars = getattr(self, "export_field_language_vars", {})
        selected_languages = []
        if isinstance(language_vars, dict):
            for code, var in language_vars.items():
                try:
                    if var.get():
                        selected_languages.append(code)
                except Exception:
                    continue
        extra_codes = getattr(self, "_export_unknown_language_codes", []) or []
        for code in extra_codes:
            if not isinstance(code, str):
                continue
            stripped = code.strip()
            if not stripped or stripped in selected_languages:
                continue
            selected_languages.append(stripped)
        field_languages = _normalize_export_field_languages(name, selected_languages)

        trimmed_template = template.strip()
        if trimmed_template:
            if _looks_like_formula(trimmed_template):
                try:
                    FormulaEngine.describe(trimmed_template)
                except FormulaError as exc:
                    show_error(f"Помилка у формулі: {exc}")
                    self._update_template_status(template)
                    return False
            else:
                try:
                    Template(trimmed_template)
                except TemplateError as exc:
                    show_error(f"Помилка у шаблоні Jinja2: {exc}")
                    self._update_template_status(template)
                    return False

        changed = False
        if field.get("field") != name:
            field["field"] = name
            changed = True
        if field.get("template", "") != template:
            field["template"] = template
            changed = True
        if bool(field.get("enabled")) != enabled:
            field["enabled"] = enabled
            changed = True
        existing_languages = []
        raw_existing_languages = field.get("languages", [])
        if isinstance(raw_existing_languages, str):
            code = raw_existing_languages.strip()
            if code:
                existing_languages.append(code)
        elif isinstance(raw_existing_languages, (list, tuple, set)):
            seen_codes = set()
            for lang in raw_existing_languages:
                if not isinstance(lang, str):
                    continue
                code = lang.strip()
                if not code or code in seen_codes:
                    continue
                existing_languages.append(code)
                seen_codes.add(code)
        if existing_languages != field_languages:
            field["languages"] = field_languages
            changed = True
        else:
            field["languages"] = existing_languages

        self._update_template_status(template)

        if changed:
            self._refresh_export_fields_tree(select_index=idx)

        if save_to_file:
            save_export_fields(self.export_fields)
            show_info("Налаштування експорту збережено.")

        return True

    def _export_save_all(self):
        self._export_apply_detail(False)
        save_export_fields(self.export_fields)
        show_info("Налаштування експорту збережено.")

    def _export_add_field(self):
        self._export_apply_detail(False)
        new_field = {"field": "Нове_поле", "template": "", "enabled": True, "languages": []}
        self.export_fields.append(new_field)
        idx = len(self.export_fields) - 1
        self._refresh_export_fields_tree(select_index=idx)
        self._load_export_field_detail(idx)

    def _export_delete_field(self):
        idx = getattr(self, "_export_selected_index", None)
        if idx is None or idx < 0 or idx >= len(self.export_fields):
            show_error("Оберіть поле для видалення.")
            return
        if not messagebox.askyesno("Підтвердження", "Видалити вибране поле?"):
            return
        self.export_fields.pop(idx)
        if self.export_fields:
            new_idx = min(idx, len(self.export_fields) - 1)
            self._refresh_export_fields_tree(select_index=new_idx)
            self._load_export_field_detail(new_idx)
        else:
            self._refresh_export_fields_tree(select_index=None)
            self._load_export_field_detail(None)

    def _export_move_field(self, direction: int):
        idx = getattr(self, "_export_selected_index", None)
        if idx is None:
            return
        new_idx = idx + direction
        if new_idx < 0 or new_idx >= len(self.export_fields):
            return
        self._export_apply_detail(False)
        self.export_fields[idx], self.export_fields[new_idx] = self.export_fields[new_idx], self.export_fields[idx]
        self._refresh_export_fields_tree(select_index=new_idx)
        self._load_export_field_detail(new_idx)

    def _export_reset_defaults(self):
        if not messagebox.askyesno("Підтвердження", "Відновити стандартний список полів?"):
            return
        self.export_fields = _copy_default_export_fields()
        save_export_fields(self.export_fields)
        if self.export_fields:
            self._refresh_export_fields_tree(select_index=0)
            self._load_export_field_detail(0)
        else:
            self._refresh_export_fields_tree(select_index=None)
            self._load_export_field_detail(None)
        show_info("Стандартні поля відновлено.")

    def _collect_selected_export_languages(self):
        vars_list = getattr(self, "export_language_vars", [])
        if not vars_list:
            return []
        selected = []
        seen = set()
        for code, var in vars_list:
            if not isinstance(code, str):
                continue
            try:
                is_checked = bool(var.get())
            except Exception:
                is_checked = False
            if is_checked:
                stripped = code.strip()
                if stripped and stripped not in seen:
                    selected.append(stripped)
                    seen.add(stripped)
        return selected

    # -------- Генерація
    def _build_tab_generate(self):
        wrap = ctk.CTkFrame(self.tab_generate)
        wrap.pack(fill="both", expand=True, padx=10, pady=10)

        left = ctk.CTkFrame(wrap)
        left.pack(side="left", fill="both", expand=True, padx=(0, 10), pady=10)

        self.gen_filter_header = ctk.CTkFrame(left)
        self.gen_filter_header.pack(fill="x", padx=10, pady=(6, 4))
        ctk.CTkLabel(
            self.gen_filter_header,
            text="Категорії / бренди / моделі",
        ).pack(side="left", padx=(0, 6))
        self.gen_filter_toggle = ctk.CTkButton(
            self.gen_filter_header,
            text="Фільтр ▴",
            width=90,
            command=self._toggle_gen_filter_panel,
        )
        self.gen_filter_toggle.pack(side="right")

        self.gen_filter_panel = ctk.CTkFrame(left)
        self.gen_filter_panel.pack(fill="x", padx=10, pady=(0, 6))
        ctk.CTkLabel(self.gen_filter_panel, text="Період:").pack(side="left", padx=(6, 6), pady=6)
        fields_frame = ctk.CTkFrame(self.gen_filter_panel)
        fields_frame.pack(side="left", fill="x", expand=True, padx=(0, 6), pady=6)

        self.gen_filter_start_date_var = tk.StringVar(value="")
        self.gen_filter_start_time_var = tk.StringVar(value="")
        self.gen_filter_end_date_var = tk.StringVar(value="")
        self.gen_filter_end_time_var = tk.StringVar(value="")

        ctk.CTkLabel(fields_frame, text="з").pack(side="left", padx=(0, 4))
        self.gen_filter_start_date_entry = ctk.CTkEntry(
            fields_frame,
            width=110,
            textvariable=self.gen_filter_start_date_var,
            placeholder_text="дд.мм.рррр",
        )
        self.gen_filter_start_date_entry.pack(side="left", padx=(0, 4))
        self.gen_filter_start_time_entry = ctk.CTkEntry(
            fields_frame,
            width=80,
            textvariable=self.gen_filter_start_time_var,
            placeholder_text="гг.хх",
        )
        self.gen_filter_start_time_entry.pack(side="left", padx=(0, 8))
        ctk.CTkLabel(fields_frame, text="по").pack(side="left", padx=(0, 4))
        self.gen_filter_end_date_entry = ctk.CTkEntry(
            fields_frame,
            width=110,
            textvariable=self.gen_filter_end_date_var,
            placeholder_text="дд.мм.рррр",
        )
        self.gen_filter_end_date_entry.pack(side="left", padx=(0, 4))
        self.gen_filter_end_time_entry = ctk.CTkEntry(
            fields_frame,
            width=80,
            textvariable=self.gen_filter_end_time_var,
            placeholder_text="гг.хх",
        )
        self.gen_filter_end_time_entry.pack(side="left", padx=(0, 4))

        for entry in (
            self.gen_filter_start_date_entry,
            self.gen_filter_start_time_entry,
            self.gen_filter_end_date_entry,
            self.gen_filter_end_time_entry,
        ):
            if entry is not None:
                entry.bind("<Return>", self._on_gen_filter_entry_submit)

        self.gen_filter_apply = ctk.CTkButton(
            self.gen_filter_panel,
            text="Застосувати",
            width=110,
            command=self._apply_gen_filter_range,
        )
        self.gen_filter_apply.pack(side="left", padx=(0, 6), pady=6)
        self.gen_filter_clear = ctk.CTkButton(
            self.gen_filter_panel,
            text="Скинути",
            width=90,
            command=self._reset_gen_filter,
        )
        self.gen_filter_clear.pack(side="left", padx=(0, 6), pady=6)
        self._gen_filter_visible = True

        tree_container = ctk.CTkFrame(left)
        tree_container.pack(fill="both", expand=True, padx=10, pady=(0, 6))
        tree_container.grid_columnconfigure(0, weight=1)
        tree_container.grid_rowconfigure(0, weight=1)

        self._gen_tree = ttk.Treeview(tree_container, show="tree", selectmode="extended")
        self._gen_tree.column("#0", anchor="w", width=320)
        self._gen_tree.grid(row=0, column=0, sticky="nsew")

        y_scroll = ttk.Scrollbar(tree_container, orient="vertical", command=self._gen_tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(tree_container, orient="horizontal", command=self._gen_tree.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")

        self._gen_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self._gen_tree.bind("<Button-1>", self._on_gen_tree_click, add="+")
        self._gen_tree.bind("<space>", self._toggle_selected_gen_node)
        self._gen_tree.bind("<Return>", self._toggle_selected_gen_node)

        controls = ctk.CTkFrame(left)
        controls.pack(fill="x", padx=10, pady=(0, 8))
        self.generate_tree_buttons = []
        btn_select_all = ctk.CTkButton(controls, text="Вибрати все", command=self._select_all_gen_tree, width=120)
        btn_select_all.pack(side="left", padx=4)
        self.generate_tree_buttons.append(btn_select_all)
        btn_clear = ctk.CTkButton(controls, text="Очистити", command=self._clear_all_gen_tree, width=120)
        btn_clear.pack(side="left", padx=4)
        self.generate_tree_buttons.append(btn_clear)
        btn_expand = ctk.CTkButton(controls, text="Розгорнути все", command=self._expand_all_gen_tree, width=140)
        btn_expand.pack(side="right", padx=4)
        self.generate_tree_buttons.append(btn_expand)
        btn_collapse = ctk.CTkButton(controls, text="Згорнути все", command=self._collapse_all_gen_tree, width=140)
        btn_collapse.pack(side="right", padx=4)
        self.generate_tree_buttons.append(btn_collapse)

        tip_text = (
            "Порада: клацніть або натисніть пробіл, щоб поставити/зняти галочку. "
            "Без вибору буде згенеровано увесь каталог."
        )
        tip = ctk.CTkLabel(left, text=tip_text, anchor="w", wraplength=320)
        tip.pack(fill="x", padx=10, pady=(0, 6))

        right = ctk.CTkFrame(wrap)
        right.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)

        fmt_frame = ctk.CTkFrame(right)
        fmt_frame.pack(fill="x", padx=10, pady=(6, 6))
        ctk.CTkLabel(fmt_frame, text="Формат експорту:").pack(anchor="w", padx=6, pady=(4, 4))
        export_formats = get_available_export_formats()
        if not export_formats:
            export_formats = ["JSON (.json)"]
        default_format = export_formats[0]
        self.export_fmt_var = tk.StringVar(value=default_format)
        self.export_fmt_menu = ctk.CTkOptionMenu(
            fmt_frame,
            values=export_formats,
            variable=self.export_fmt_var,
            width=200,
        )
        self.export_fmt_menu.pack(anchor="w", padx=6, pady=(0, 4))

        path_frame = ctk.CTkFrame(right)
        path_frame.pack(fill="x", padx=10, pady=(4, 6))
        ctk.CTkLabel(path_frame, text="Папка збереження:").pack(anchor="w", padx=6, pady=(4, 4))
        default_export_dir = self.settings.get("export_folder") or str(get_default_export_dir())
        self.out_folder_var = tk.StringVar(value=default_export_dir)
        self.out_folder_entry = ctk.CTkEntry(path_frame, textvariable=self.out_folder_var)
        self.out_folder_entry.pack(fill="x", padx=6, pady=(0, 4))
        self._bind_clipboard_shortcuts(self.out_folder_entry)
        self.theme_manager.register(self.out_folder_entry, "widget")
        self.choose_folder_button = ctk.CTkButton(
            path_frame,
            text="Обрати...",
            command=self._choose_folder,
            width=110,
        )
        self.choose_folder_button.pack(anchor="e", padx=6, pady=(0, 4))
        self.theme_manager.register(self.choose_folder_button, "accent_button")

        languages_frame = ctk.CTkFrame(right)
        languages_frame.pack(fill="x", padx=10, pady=(4, 6))
        ctk.CTkLabel(languages_frame, text="Мови експорту:").pack(anchor="w", padx=6, pady=(4, 2))
        lang_checks = ctk.CTkFrame(languages_frame, fg_color="transparent")
        lang_checks.pack(fill="x", padx=6, pady=(2, 4))
        self.generate_language_checks_container = lang_checks
        self.generate_language_hint = ctk.CTkLabel(
            languages_frame,
            text="",
            anchor="w",
            justify="left",
            wraplength=360,
            font=ctk.CTkFont(size=12),
        )
        self.generate_language_hint.pack(fill="x", padx=6, pady=(0, 2))
        self._build_generate_language_checkboxes()

        types_frame = ctk.CTkFrame(right)
        types_frame.pack(fill="both", expand=True, padx=10, pady=(4, 6))
        ctk.CTkLabel(types_frame, text="Типи плівок:").pack(anchor="w", padx=6, pady=(4, 2))
        types_wrap = ctk.CTkFrame(types_frame)
        types_wrap.pack(fill="x", padx=6, pady=(2, 6))
        types_scroll = ctk.CTkScrollableFrame(types_wrap, height=160)
        types_scroll.pack(fill="both", expand=True, padx=8, pady=8)
        types_scroll.bind("<Configure>", self._on_film_types_resize)
        self._film_types_scroll = types_scroll
        self.ft_vars = []

        action_row = ctk.CTkFrame(right)
        action_row.pack(fill="x", padx=10, pady=(6, 0))
        self.generate_preview_button = ctk.CTkButton(
            action_row,
            text="Попередній перегляд",
            command=self._preview_generation,
            height=36,
        )
        self.generate_preview_button.pack(side="right", padx=6)
        self.theme_manager.register(self.generate_preview_button, "accent_button")
        self.generate_run_button = ctk.CTkButton(
            action_row,
            text="Згенерувати",
            command=self._generate,
            height=36,
        )
        self.generate_run_button.pack(side="right", padx=6)
        self.theme_manager.register(self.generate_run_button, "accent_button")

        progress_frame = ctk.CTkFrame(right)
        progress_frame.pack(fill="x", padx=10, pady=(10, 0))
        self.progress_bar = ctk.CTkProgressBar(progress_frame)
        self.progress_bar.pack(fill="x", padx=6, pady=(6, 4))
        self.progress_bar.set(0)
        self.progress_label = ctk.CTkLabel(progress_frame, text="Очікування", anchor="w")
        self.progress_label.pack(fill="x", padx=6, pady=(0, 6))

        self._reload_gen_tree()

    def _toggle_gen_filter_panel(self):
        panel = getattr(self, "gen_filter_panel", None)
        header = getattr(self, "gen_filter_header", None)
        button = getattr(self, "gen_filter_toggle", None)
        if panel is None or button is None or header is None:
            return
        if getattr(self, "_gen_filter_visible", False):
            panel.pack_forget()
            self._gen_filter_visible = False
            button.configure(text="Фільтр ▾")
        else:
            panel.pack(fill="x", padx=10, pady=(0, 6), after=header)
            self._gen_filter_visible = True
            button.configure(text="Фільтр ▴")

    def _reset_gen_filter(self):
        for var in (
            getattr(self, "gen_filter_start_date_var", None),
            getattr(self, "gen_filter_start_time_var", None),
            getattr(self, "gen_filter_end_date_var", None),
            getattr(self, "gen_filter_end_time_var", None),
        ):
            if isinstance(var, tk.StringVar):
                var.set("")
        self._gen_filter_range = (None, None)
        self._reload_gen_tree()

    def _on_gen_filter_entry_submit(self, _event=None):
        self._apply_gen_filter_range()

    def _parse_gen_filter_datetime(self, date_var, time_var, label: str) -> Optional[datetime]:
        date_text = ""
        time_text = ""
        if isinstance(date_var, tk.StringVar):
            date_text = date_var.get().strip()
        if isinstance(time_var, tk.StringVar):
            time_text = time_var.get().strip()
        if not date_text and not time_text:
            return None
        if not date_text or not time_text:
            raise ValueError(f"Вкажіть повну дату і час для поля '{label}'.")
        normalized_time = time_text.replace(" ", "")
        if ":" not in normalized_time and "." in normalized_time:
            normalized_time = normalized_time.replace(".", ":")
        candidate = f"{date_text} {normalized_time}"
        for fmt in ("%d.%m.%Y %H:%M", "%d.%m.%Y %H.%M"):
            try:
                return datetime.strptime(candidate, fmt)
            except ValueError:
                continue
        raise ValueError(
            f"Невірний формат для поля '{label}'. Використовуйте 'дд.мм.рррр' та 'гг.хх'."
        )

    def _apply_gen_filter_range(self):
        try:
            start = self._parse_gen_filter_datetime(
                getattr(self, "gen_filter_start_date_var", None),
                getattr(self, "gen_filter_start_time_var", None),
                "Початок",
            )
            end = self._parse_gen_filter_datetime(
                getattr(self, "gen_filter_end_date_var", None),
                getattr(self, "gen_filter_end_time_var", None),
                "Кінець",
            )
        except ValueError as exc:
            show_error(str(exc))
            return

        if start is not None and end is not None and start > end:
            show_error("Дата початку не може бути пізнішою за дату завершення.")
            return

        self._gen_filter_range = (start, end)
        self._reload_gen_tree()

    def _get_gen_filter_range(self) -> Tuple[Optional[datetime], Optional[datetime]]:
        value = getattr(self, "_gen_filter_range", (None, None))
        if isinstance(value, tuple) and len(value) == 2:
            return cast(Tuple[Optional[datetime], Optional[datetime]], value)
        return (None, None)

    def _parse_db_timestamp(self, raw: Optional[str]) -> Optional[datetime]:
        if raw is None:
            return None
        if isinstance(raw, bytes):
            raw = raw.decode("utf-8", errors="ignore")
        if not isinstance(raw, str):
            raw = str(raw)
        text = raw.strip()
        if not text:
            return None
        try:
            return datetime.fromisoformat(text)
        except ValueError:
            pass
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
            try:
                return datetime.strptime(text, fmt)
            except ValueError:
                continue
        return None

    def _gen_filter_matches(
        self,
        created_at: Optional[str],
        date_range: Tuple[Optional[datetime], Optional[datetime]],
    ) -> bool:
        start, end = date_range if date_range else (None, None)
        if start is None and end is None:
            return True
        created = self._parse_db_timestamp(created_at)
        if created is None:
            return False
        if start is not None and created < start:
            return False
        if end is not None and created > end:
            return False
        return True

    def _queue_progress_update(self, current: int, total: int, stage: str) -> None:
        self._ensure_background_primitives()
        if not hasattr(self, "_progress_lock"):
            return
        with self._progress_lock:
            now = time.time()
            last = getattr(self, "_last_progress_update", 0.0)
            if current not in (0, total) and now - last < 0.05:
                return
            self._last_progress_update = now
        self._call_in_ui_thread(self._progress_update, current, total, stage=stage)

    def _queue_progress_message(self, message: str) -> None:
        self._call_in_ui_thread(self._progress_message, message)

    def _queue_generation_error(self, message: str, reset_label: str = "Очікування") -> None:
        def _handle() -> None:
            self._progress_reset(reset_label)
            self._finalize_generation_task()
            self._schedule_progress_idle()
            show_error(message)

        self._call_in_ui_thread(_handle)

    def _finalize_generation_task(self) -> None:
        self._ensure_background_primitives()
        self._generation_task_running = False
        self._active_generation_thread = None
        with self._progress_lock:
            self._last_progress_update = 0.0
        self._set_generate_controls_state(True)

    def _schedule_progress_idle(self, delay_ms: int = 1800) -> None:
        try:
            self.after(delay_ms, lambda: self._progress_reset("Очікування"))
        except Exception:
            pass

    def _progress_reset(self, message: str = "Очікування"):
        bar = getattr(self, "progress_bar", None)
        if bar is None:
            return
        if hasattr(bar, "stop"):
            bar.stop()
        bar.configure(mode="determinate")
        bar.set(0)
        label = getattr(self, "progress_label", None)
        if label is not None and message is not None:
            label.configure(text=message)
        self.update_idletasks()

    def _progress_update(self, current: int, total: int, stage: str = "Генерація"):
        bar = getattr(self, "progress_bar", None)
        if bar is None:
            return
        total = total or 0
        if total > 0:
            fraction = max(0.0, min(float(current) / float(total), 1.0))
        else:
            fraction = 0.0
        bar.configure(mode="determinate")
        bar.set(fraction)
        label = getattr(self, "progress_label", None)
        if label is not None:
            if total > 0:
                text = f"{stage}: {current} з {total}"
            else:
                text = f"{stage}: {current}"
            label.configure(text=text)
        self.update_idletasks()

    def _progress_message(self, message: str):
        label = getattr(self, "progress_label", None)
        if label is not None and message is not None:
            label.configure(text=message)
        self.update_idletasks()

    def _progress_finish(self, message: str = "Готово"):
        bar = getattr(self, "progress_bar", None)
        if bar is not None:
            if hasattr(bar, "stop"):
                bar.stop()
            bar.configure(mode="determinate")
            bar.set(1)
        label = getattr(self, "progress_label", None)
        if label is not None and message is not None:
            label.configure(text=message)
        self.update_idletasks()

    def _reload_gen_tree(self):
        tree = getattr(self, "_gen_tree", None)
        if tree is None:
            return

        prev_checked = self._collect_checked_model_ids()
        prev_open = set()

        def _gather_open(iid):
            if tree.item(iid, "open"):
                prev_open.add(iid)
            for child in tree.get_children(iid):
                _gather_open(child)

        for root_iid in tree.get_children(""):
            _gather_open(root_iid)

        tree.delete(*tree.get_children(""))
        self._gen_tree_states.clear()
        self._gen_tree_meta.clear()
        self._gen_tree_labels.clear()

        def _clean_label(value):
            if isinstance(value, str):
                return value.strip()
            if value is None:
                return ""
            return str(value)

        def _split_meta(row):
            if not row:
                return None, None, None
            rid = row[0]
            name = row[1] if len(row) > 1 else None
            created = row[2] if len(row) > 2 else None
            if isinstance(name, str):
                name = name.strip()
            if isinstance(created, str):
                created = created.strip()
            return rid, name, created

        filter_range = self._get_gen_filter_range()
        range_active = False
        if isinstance(filter_range, tuple) and len(filter_range) == 2:
            range_active = bool(filter_range[0] or filter_range[1])

        for cat_row in get_categories(include_created=True):
            cat_id, cat_name, cat_created = _split_meta(cat_row)
            if cat_id is None:
                continue
            label = _clean_label(cat_name)
            cat_matches = self._gen_filter_matches(cat_created, filter_range)
            brand_nodes = []

            for brand_row in get_brands(cat_id, include_created=True):
                brand_id, brand_name, brand_created = _split_meta(brand_row)
                if brand_id is None:
                    continue
                models_raw = []
                for model_row in get_models(brand_id, include_created=True):
                    model_id, model_name, model_created = _split_meta(model_row)
                    if model_id is None:
                        continue
                    models_raw.append((model_id, model_name, model_created))
                if not range_active:
                    models_to_show = models_raw
                else:
                    models_to_show = [
                        (mid, mname, mcreated)
                        for mid, mname, mcreated in models_raw
                        if self._gen_filter_matches(mcreated, filter_range)
                    ]
                brand_matches = self._gen_filter_matches(brand_created, filter_range)
                if not models_to_show and not brand_matches:
                    continue
                brand_nodes.append((brand_id, brand_name, brand_created, models_to_show))

            if not brand_nodes and not cat_matches:
                continue

            cat_iid = f"cat_{cat_id}"
            self._gen_tree_labels[cat_iid] = label
            self._gen_tree_states[cat_iid] = 0
            self._gen_tree_meta[cat_iid] = {
                "type": "category",
                "id": cat_id,
                "created_at": cat_created,
            }
            tree.insert("", "end", iid=cat_iid, text=f"{self._state_symbol(0)} {label}")

            if not brand_nodes:
                continue

            for brand_id, brand_name, brand_created, models_to_show in brand_nodes:
                b_label = _clean_label(brand_name)
                brand_iid = f"brand_{brand_id}"
                self._gen_tree_labels[brand_iid] = b_label
                self._gen_tree_states[brand_iid] = 0
                self._gen_tree_meta[brand_iid] = {
                    "type": "brand",
                    "id": brand_id,
                    "category_id": cat_id,
                    "created_at": brand_created,
                }
                tree.insert(cat_iid, "end", iid=brand_iid, text=f"{self._state_symbol(0)} {b_label}")

                for model_id, model_name, model_created in models_to_show:
                    m_label = _clean_label(model_name)
                    model_iid = f"model_{model_id}"
                    self._gen_tree_labels[model_iid] = m_label
                    self._gen_tree_states[model_iid] = 0
                    self._gen_tree_meta[model_iid] = {
                        "type": "model",
                        "id": model_id,
                        "brand_id": brand_id,
                        "category_id": cat_id,
                        "created_at": model_created,
                    }
                    tree.insert(brand_iid, "end", iid=model_iid, text=f"{self._state_symbol(0)} {m_label}")

        if not prev_open:
            roots = tree.get_children("")
            if roots:
                tree.item(roots[0], open=True)
        else:
            for iid in prev_open:
                if tree.exists(iid):
                    tree.item(iid, open=True)

        for mid in sorted(prev_checked):
            iid = f"model_{mid}"
            if tree.exists(iid):
                self._set_gen_tree_state(iid, 2, propagate=False)
                self._update_parent_states(iid)

    def _state_symbol(self, state: int) -> str:
        if state == 2:
            return "☑"
        if state == 1:
            return "◪"
        return "☐"

    def _set_gen_tree_state(self, iid: str, state: int, propagate: bool = False):
        tree = getattr(self, "_gen_tree", None)
        if tree is None or not tree.exists(iid):
            return
        self._gen_tree_states[iid] = state
        label = self._gen_tree_labels.get(iid, tree.item(iid, "text"))
        display = f"{self._state_symbol(state)} {label}"
        tree.item(iid, text=display)
        if propagate:
            for child in tree.get_children(iid):
                self._set_gen_tree_state(child, state, propagate=True)

    def _update_parent_states(self, iid: str):
        tree = getattr(self, "_gen_tree", None)
        if tree is None:
            return
        parent = tree.parent(iid)
        if not parent:
            return
        child_states = [self._gen_tree_states.get(child, 0) for child in tree.get_children(parent)]
        if all(state == 2 for state in child_states):
            new_state = 2
        elif all(state == 0 for state in child_states):
            new_state = 0
        else:
            new_state = 1
        self._set_gen_tree_state(parent, new_state, propagate=False)
        self._update_parent_states(parent)

    def _collect_checked_model_ids(self):
        ids = set()
        for iid, state in self._gen_tree_states.items():
            if state != 2:
                continue
            meta = self._gen_tree_meta.get(iid)
            if meta and meta.get("type") == "model":
                mid = meta.get("id")
                if mid:
                    try:
                        ids.add(int(mid))
                    except (TypeError, ValueError):
                        continue
        return ids

    def _on_gen_tree_click(self, event):
        tree = getattr(self, "_gen_tree", None)
        if tree is None:
            return
        iid = tree.identify_row(event.y)
        if not iid:
            return
        element = tree.identify_element(event.x, event.y)
        if element != "text":
            return
        self._toggle_gen_tree_node(iid)

    def _toggle_selected_gen_node(self, event=None):
        tree = getattr(self, "_gen_tree", None)
        if tree is None:
            return "break"
        iid = tree.focus()
        if not iid:
            selection = tree.selection()
            if selection:
                iid = selection[0]
        if not iid:
            return "break"
        self._toggle_gen_tree_node(iid)
        return "break"

    def _toggle_gen_tree_node(self, iid: str):
        if iid not in self._gen_tree_states:
            return
        current = self._gen_tree_states.get(iid, 0)
        new_state = 0 if current == 2 else 2
        self._set_gen_tree_state(iid, new_state, propagate=True)
        self._update_parent_states(iid)

    def _set_gen_tree_open_recursive(self, iid: str, value: bool):
        tree = getattr(self, "_gen_tree", None)
        if tree is None or not tree.exists(iid):
            return
        tree.item(iid, open=value)
        for child in tree.get_children(iid):
            self._set_gen_tree_open_recursive(child, value)

    def _expand_all_gen_tree(self):
        tree = getattr(self, "_gen_tree", None)
        if tree is None:
            return
        for iid in tree.get_children(""):
            self._set_gen_tree_open_recursive(iid, True)

    def _collapse_all_gen_tree(self):
        tree = getattr(self, "_gen_tree", None)
        if tree is None:
            return
        for iid in tree.get_children(""):
            self._set_gen_tree_open_recursive(iid, False)

    def _select_all_gen_tree(self):
        tree = getattr(self, "_gen_tree", None)
        if tree is None:
            return
        for iid in tree.get_children(""):
            self._set_gen_tree_state(iid, 2, propagate=True)

    def _clear_all_gen_tree(self):
        tree = getattr(self, "_gen_tree", None)
        if tree is None:
            return
        for iid in tree.get_children(""):
            self._set_gen_tree_state(iid, 0, propagate=True)

    def _set_generate_controls_state(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        widgets = [
            getattr(self, "generate_preview_button", None),
            getattr(self, "generate_run_button", None),
            getattr(self, "gen_filter_apply", None),
            getattr(self, "gen_filter_clear", None),
            getattr(self, "gen_filter_toggle", None),
            getattr(self, "choose_folder_button", None),
            getattr(self, "export_fmt_menu", None),
        ]
        for widget in widgets:
            if widget is None:
                continue
            try:
                widget.configure(state=state)
            except Exception:
                pass

        entries = [
            getattr(self, "gen_filter_start_date_entry", None),
            getattr(self, "gen_filter_start_time_entry", None),
            getattr(self, "gen_filter_end_date_entry", None),
            getattr(self, "gen_filter_end_time_entry", None),
            getattr(self, "out_folder_entry", None),
        ]
        for entry in entries:
            if entry is None:
                continue
            try:
                entry.configure(state=state)
            except Exception:
                pass

        for button in getattr(self, "generate_tree_buttons", []):
            try:
                button.configure(state=state)
            except Exception:
                continue

        for filmtype in getattr(self, "ft_vars", []):
            widget = None
            if isinstance(filmtype, dict):
                widget = filmtype.get("widget")
            elif isinstance(filmtype, (list, tuple)) and len(filmtype) >= 3:
                widget = filmtype[2]
            if widget is None:
                continue
            try:
                widget.configure(state=state)
            except Exception:
                continue

        language_container = getattr(self, "generate_language_checks_container", None)
        if language_container is not None:
            for child in language_container.winfo_children():
                try:
                    child.configure(state=state)
                except Exception:
                    continue

        tree = getattr(self, "_gen_tree", None)
        if tree is not None:
            try:
                tree.configure(selectmode="extended" if enabled else "none")
            except Exception:
                pass

    def _compute_film_type_cols(self, width_px: int) -> int:
        col_w = 240
        cols = max(1, width_px // col_w)
        return min(cols, 6)

    def _build_film_type_checkboxes(self, parent, film_types: List[Dict[str, object]]) -> None:
        self._film_type_order = []
        for item in film_types:
            if not isinstance(item, dict):
                continue
            name = item.get("name")
            if not isinstance(name, str) or not name.strip():
                continue
            normalized_name = name.strip()
            self._film_type_order.append(normalized_name)
            var = self._film_type_vars.get(normalized_name)
            enabled = bool(item.get("enabled", True))
            if var is None:
                var = tk.BooleanVar(value=enabled)
                self._film_type_vars[normalized_name] = var
            else:
                var.set(enabled)
            if normalized_name not in self._film_type_cbs:
                checkbox = ctk.CTkCheckBox(parent, text=normalized_name, variable=var)
                self._film_type_cbs[normalized_name] = checkbox

        for name in list(self._film_type_cbs.keys()):
            if name not in self._film_type_order:
                widget = self._film_type_cbs.pop(name)
                widget.destroy()
                self._film_type_vars.pop(name, None)

        self.ft_vars = [
            {"name": name, "var": self._film_type_vars[name], "widget": self._film_type_cbs[name]}
            for name in self._film_type_order
        ]

    def _relayout_film_types(self) -> None:
        parent = self._film_types_scroll
        if parent is None:
            return
        width = parent.winfo_width()
        if width < 50:
            return
        cols = self._compute_film_type_cols(width)
        if cols == self._film_types_cols:
            return
        self._film_types_cols = cols
        for cb in self._film_type_cbs.values():
            cb.grid_forget()
        for c in range(cols):
            parent.grid_columnconfigure(c, weight=1)
        for i, name in enumerate(self._film_type_order):
            cb = self._film_type_cbs.get(name)
            if cb is None:
                continue
            r = i // cols
            c = i % cols
            cb.grid(row=r, column=c, sticky="w", padx=10, pady=6)

    def _schedule_film_types_relayout(self) -> None:
        if self._film_layout_job:
            try:
                self.after_cancel(self._film_layout_job)
            except Exception:
                pass
        self._film_layout_job = self.after(80, self._relayout_film_types)

    def _on_film_types_resize(self, _event=None) -> None:
        self._schedule_film_types_relayout()

    def _refresh_filmtype_checkboxes(self):
        frame = self._film_types_scroll
        if frame is None:
            return
        film_types = self.templates.get("film_types", []) if isinstance(self.templates, dict) else []
        self._build_film_type_checkboxes(frame, film_types)
        self._schedule_film_types_relayout()
        self._refresh_template_selectors()

    def _choose_folder(self):
        folder = filedialog.askdirectory(title="Виберіть папку для файлів")
        if folder:
            self.out_folder_var.set(folder)
            self.settings["export_folder"] = folder
            save_settings(self.settings)

    def _collect_generation_context(self) -> Optional[Dict[str, object]]:
        entries = getattr(self, "ft_vars", [])
        selected_types = []
        for entry in entries:
            name = None
            var = None
            if isinstance(entry, dict):
                name = entry.get("name")
                var = entry.get("var")
            elif isinstance(entry, (list, tuple)) and len(entry) >= 2:
                name, var = entry[0], entry[1]
            if isinstance(name, str) and hasattr(var, "get") and callable(getattr(var, "get")):
                try:
                    if var.get():
                        selected_types.append(name)
                except Exception:
                    continue
        if not selected_types:
            show_error("Оберіть хоча б один тип плівки.")
            return None

        film_types_store = self.templates.get("film_types") if isinstance(self.templates, dict) else None
        if isinstance(film_types_store, list):
            for entry in entries:
                name = None
                var = None
                if isinstance(entry, dict):
                    name = entry.get("name")
                    var = entry.get("var")
                elif isinstance(entry, (list, tuple)) and len(entry) >= 2:
                    name, var = entry[0], entry[1]
                if not isinstance(name, str) or not hasattr(var, "get"):
                    continue
                for item in film_types_store:
                    if isinstance(item, dict) and item.get("name") == name:
                        try:
                            item["enabled"] = bool(var.get())
                        except Exception:
                            item["enabled"] = True
                        break
        save_templates(self.templates)

        selected_models = sorted(self._collect_checked_model_ids())
        selected_languages = self._collect_selected_export_languages()
        if self._template_language_codes() and self.export_language_vars and not selected_languages:
            show_error("Оберіть хоча б одну мову експорту.")
            return None

        export_format = self.export_fmt_var.get() if hasattr(self, "export_fmt_var") else "JSON (.json)"
        default_export_dir = str(get_default_export_dir())
        output_folder_raw = self.out_folder_var.get().strip() if hasattr(self, "out_folder_var") else default_export_dir
        if not output_folder_raw:
            output_folder = default_export_dir
        else:
            candidate = Path(output_folder_raw).expanduser()
            if candidate.is_dir():
                output_folder = str(candidate)
            else:
                output_folder = default_export_dir
        if hasattr(self, "out_folder_var"):
            try:
                self.out_folder_var.set(output_folder)
            except Exception:
                pass

        context: Dict[str, object] = {
            "film_types": list(selected_types),
            "templates": deepcopy(self.templates),
            "title_tags": deepcopy(self.title_tags_templates),
            "export_fields": deepcopy(self.export_fields),
            "selected_models": list(selected_models),
            "selected_languages": list(selected_languages),
            "export_format": export_format,
            "output_folder": output_folder,
        }
        return context

    def _on_preview_ready(
        self,
        columns: Sequence[str],
        preview_records: Sequence[Sequence[str]],
        total_count: int,
    ) -> None:
        self._progress_finish(f"Попередній перегляд: {total_count} рядків")
        self._finalize_generation_task()
        self._show_preview_window(columns, preview_records, total_count)
        self._schedule_progress_idle()

    def _show_preview_window(
        self,
        columns: Sequence[str],
        preview_records: Sequence[Sequence[str]],
        total_count: int,
    ) -> None:
        existing = getattr(self, "_preview_window", None)
        if existing is not None:
            try:
                if existing.winfo_exists():
                    existing.destroy()
            except Exception:
                pass

        preview_window = ctk.CTkToplevel(self)
        preview_window.title("Попередній перегляд генерації")
        preview_window.geometry("960x480")
        self._preview_window = preview_window

        info_text = f"Показано перші {len(preview_records)} з {total_count} рядків."
        ctk.CTkLabel(preview_window, text=info_text, anchor="w").pack(fill="x", padx=14, pady=(12, 4))

        table_frame = ctk.CTkFrame(preview_window)
        table_frame.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        column_ids = [f"preview_col_{idx}" for idx in range(len(columns))]
        tree = ttk.Treeview(table_frame, columns=column_ids, show="headings")
        tree.grid(row=0, column=0, sticky="nsew")

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        for col_id, header in zip(column_ids, columns):
            tree.heading(col_id, text=header)
            tree.column(col_id, anchor="w", stretch=True, width=160)

        tree.tag_configure("odd", background="#20242b")
        tree.tag_configure("even", background="#151921")

        for idx, record in enumerate(preview_records):
            values = _row_to_values(record, columns)
            tag = "odd" if idx % 2 else "even"
            tree.insert("", "end", values=values, tags=(tag,))

        if total_count > len(preview_records):
            note = f"(Доступно більше рядків: всього {total_count}.)"
            ctk.CTkLabel(preview_window, text=note, anchor="w").pack(fill="x", padx=14, pady=(0, 10))

        ctk.CTkButton(preview_window, text="Закрити", command=preview_window.destroy).pack(pady=(0, 12))

    def _preview_generation(self):
        self._ensure_background_primitives()
        if getattr(self, "_generation_task_running", False):
            return show_info("Інший процес вже виконується. Зачекайте його завершення.")

        self._save_title_tags(show_message=False)
        self._export_apply_detail(save_to_file=False)
        context = self._collect_generation_context()
        if context is None:
            return

        self._generation_task_running = True
        self._set_generate_controls_state(False)
        self._progress_reset("Готуємо попередній перегляд...")
        self._progress_message("Генеруємо попередній перегляд…")
        with self._progress_lock:
            self._last_progress_update = 0.0

        preview_limit = 20

        def worker() -> None:
            try:
                def progress_callback(current: int, total: int) -> None:
                    self._queue_progress_update(current, total, stage="Попередній перегляд")

                extra_kwargs: Dict[str, object] = {
                    "languages": context.get("selected_languages"),
                    "progress_callback": progress_callback,
                }
                if context.get("selected_models"):
                    extra_kwargs["model_ids"] = context.get("selected_models")
                records, columns = generate_export_rows(
                    context.get("film_types", []),
                    context.get("templates", {}),
                    context.get("title_tags", {}),
                    context.get("export_fields", []),
                    **extra_kwargs,
                )
            except ValueError as err:
                logger.info("Помилка під час формування попереднього перегляду: %s", err)
                self._queue_generation_error(str(err), reset_label="Помилка попереднього перегляду")
                return
            except Exception as exc:
                logger.exception("Unexpected error during preview generation")
                self._queue_generation_error(
                    f"Неочікувана помилка: {exc}",
                    reset_label="Помилка попереднього перегляду",
                )
                return

            if not records:
                self._queue_generation_error(
                    "Немає даних для генерації (перевірте моделі).",
                    reset_label="Очікування",
                )
                return

            total_count = len(records)
            preview_records = records[:preview_limit]
            self._call_in_ui_thread(lambda: self._on_preview_ready(columns, preview_records, total_count))

        self._active_generation_thread = self._start_background_task(worker, name="preview-generation")

    def _generate(self):
        self._ensure_background_primitives()
        if getattr(self, "_generation_task_running", False):
            show_info("Інший процес вже виконується. Дочекайтеся завершення попередньої операції.")
            return

        self._progress_reset("Підготовка...")
        self._save_title_tags(show_message=False)
        self._export_apply_detail(save_to_file=False)
        context = self._collect_generation_context()
        if context is None:
            self._progress_reset("Очікування")
            return

        self._generation_task_running = True
        self._set_generate_controls_state(False)
        self._progress_message("Генерація даних...")
        with self._progress_lock:
            self._last_progress_update = 0.0

        def worker() -> None:
            try:
                def progress_callback(current: int, total: int) -> None:
                    self._queue_progress_update(current, total, stage="Генерація")

                extra_kwargs: Dict[str, object] = {
                    "languages": context.get("selected_languages"),
                    "progress_callback": progress_callback,
                }
                if context.get("selected_models"):
                    extra_kwargs["model_ids"] = context.get("selected_models")
                records, columns = generate_export_rows(
                    context.get("film_types", []),
                    context.get("templates", {}),
                    context.get("title_tags", {}),
                    context.get("export_fields", []),
                    **extra_kwargs,
                )
            except ValueError as err:
                logger.info("Помилка під час генерації: %s", err)
                self._queue_generation_error(str(err), reset_label="Помилка генерації")
                return
            except Exception as exc:
                logger.exception("Unexpected error during generation")
                self._queue_generation_error(
                    f"Неочікувана помилка: {exc}",
                    reset_label="Помилка генерації",
                )
                return

            if not records:
                self._queue_generation_error(
                    "Немає даних для генерації (перевірте моделі).",
                    reset_label="Очікування",
                )
                return

            self._queue_progress_message("Експорт файлів...")
            try:
                products_file = export_products(
                    records,
                    columns,
                    context.get("export_format", "JSON (.json)"),
                    context.get("output_folder", os.getcwd()),
                )
            except ExportError as exc:
                self._queue_generation_error(
                    f"Помилка експорту (код {exc.code}): {exc.message}",
                    reset_label="Помилка експорту",
                )
                return
            except Exception as exc:
                logger.exception("Unexpected error during export")
                self._queue_generation_error(
                    f"Не вдалося зберегти файли: {exc}",
                    reset_label="Помилка експорту",
                )
                return

            count = len(records)
            self._call_in_ui_thread(lambda: self._on_generation_success(count, products_file))

        self._active_generation_thread = self._start_background_task(worker, name="generate-products")

    def _on_generation_success(self, row_count: int, products_file: str) -> None:
        self._progress_finish(f"Готово: {row_count} рядків")
        self._finalize_generation_task()
        self._schedule_progress_idle()
        if products_file:
            show_info(f"✅ Згенеровано {row_count} рядків.\nФайл експорту: {products_file}")
