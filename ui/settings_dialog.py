"""Settings dialog (Total Commander style)."""
from __future__ import annotations

from copy import deepcopy
from typing import Any, Callable, Dict, Optional

import tkinter as tk
import tkinter.font as tkfont
from tkinter import colorchooser, filedialog, messagebox, ttk
import logging

import customtkinter as ctk

from app_paths import get_data_dir
from settings_service import default_settings, normalize_hex_color

logger = logging.getLogger(__name__)


class SettingsDialog(ctk.CTkToplevel):
    def __init__(
        self,
        parent: ctk.CTk,
        current_settings: Dict[str, Any],
        on_apply: Optional[Callable[[Dict[str, Any]], None]] = None,
    ) -> None:
        super().__init__(parent)
        self.title("Налаштування")
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        w = min(900, int(screen_w * 0.90))
        h = min(700, int(screen_h * 0.85))
        x = (screen_w - w) // 2
        y = (screen_h - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(720, 420)
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()

        self._parent = parent
        self._on_apply = on_apply
        self._default_settings = default_settings()
        self._draft_settings = deepcopy(current_settings)
        self._color_errors: set[str] = set()
        self._color_entries: Dict[str, ctk.CTkEntry] = {}
        self._color_swatches: Dict[str, ctk.CTkFrame] = {}
        self._color_error_labels: Dict[str, ctk.CTkLabel] = {}
        self._entry_border_colors: Dict[str, str] = {}
        self._colors_scroll: Optional[ctk.CTkScrollableFrame] = None
        self._fonts_scroll: Optional[ctk.CTkScrollableFrame] = None

        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self.bind("<Escape>", self._on_cancel_event)
        self.bind("<Control-s>", self._on_apply_shortcut)
        self.bind("<Control-S>", self._on_apply_shortcut)
        self.bind("<Return>", self._on_ok_shortcut)

        self._build_ui()
        self._select_category("Загальні")

    def _build_ui(self) -> None:
        container = ctk.CTkFrame(self)
        container.pack(fill="both", expand=True, padx=12, pady=12)
        container.grid_columnconfigure(1, weight=1)
        container.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(container, width=190)
        left.grid(row=0, column=0, sticky="ns", padx=(0, 10), pady=8)

        right = ctk.CTkFrame(container)
        right.grid(row=0, column=1, sticky="nsew", padx=(0, 8), pady=8)

        self._category_list = tk.Listbox(
            left,
            height=6,
            exportselection=False,
            activestyle="none",
            highlightthickness=1,
            borderwidth=0,
        )
        self._category_list.pack(fill="both", expand=True, padx=8, pady=8)
        categories = ["Загальні", "Оформлення", "Кольори", "Шрифти"]
        for item in categories:
            self._category_list.insert(tk.END, item)
        self._category_list.bind("<<ListboxSelect>>", self._on_category_select)
        if categories:
            self._category_list.selection_set(0)
            self._category_list.activate(0)

        self._panels: Dict[str, ctk.CTkFrame] = {}
        self._build_general_panel(right)
        self._build_appearance_panel(right)
        self._build_colors_panel(right)
        self._build_fonts_panel(right)

        footer = ctk.CTkFrame(self)
        footer.pack(fill="x", padx=12, pady=(0, 12))
        footer.grid_columnconfigure(0, weight=1)
        footer.grid_columnconfigure(1, weight=0)

        reset_btn = ctk.CTkButton(footer, text="Скинути все", command=self._reset_all)
        reset_btn.grid(row=0, column=0, sticky="w", padx=6, pady=6)

        button_box = ctk.CTkFrame(footer, fg_color="transparent")
        button_box.grid(row=0, column=1, sticky="e", padx=6, pady=6)

        button_width = 100
        self._ok_btn = ctk.CTkButton(button_box, text="OK", width=button_width, command=self._on_ok)
        self._ok_btn.pack(side="left", padx=4)
        cancel_btn = ctk.CTkButton(button_box, text="Cancel", width=button_width, command=self._on_cancel)
        cancel_btn.pack(side="left", padx=4)
        self._apply_btn = ctk.CTkButton(button_box, text="Apply", width=button_width, command=self._on_apply_click)
        self._apply_btn.pack(side="left", padx=4)

    def _build_general_panel(self, parent: ctk.CTkFrame) -> None:
        panel = ctk.CTkFrame(parent)
        panel.pack(fill="both", expand=True)
        self._panels["Загальні"] = panel

        info_frame = ctk.CTkFrame(panel)
        info_frame.pack(fill="x", padx=12, pady=12)
        info_frame.grid_columnconfigure(0, weight=1)

        data_dir = str(get_data_dir())
        ctk.CTkLabel(info_frame, text="Каталог даних:").grid(row=0, column=0, sticky="w", padx=12, pady=(12, 4))
        data_entry = ctk.CTkEntry(info_frame)
        data_entry.insert(0, data_dir)
        data_entry.configure(state="readonly")
        data_entry.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 12))

        ctk.CTkLabel(info_frame, text="Папка експорту (за замовчуванням):").grid(
            row=2,
            column=0,
            sticky="w",
            padx=12,
            pady=(0, 4),
        )
        self._export_folder_var = tk.StringVar(value=self._draft_settings.get("export_folder", ""))
        export_entry = ctk.CTkEntry(info_frame, textvariable=self._export_folder_var)
        export_entry.grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 4))

        browse_btn = ctk.CTkButton(info_frame, text="Обрати...", command=self._choose_export_folder)
        browse_btn.grid(row=4, column=0, sticky="e", padx=12, pady=(0, 12))

    def _build_appearance_panel(self, parent: ctk.CTkFrame) -> None:
        panel = ctk.CTkFrame(parent)
        panel.pack(fill="both", expand=True)
        self._panels["Оформлення"] = panel

        appearance_frame = ctk.CTkFrame(panel)
        appearance_frame.pack(fill="x", padx=12, pady=12)
        appearance_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(appearance_frame, text="Режим оформлення:").grid(
            row=0,
            column=0,
            sticky="w",
            padx=12,
            pady=(12, 4),
        )
        self._appearance_var = tk.StringVar(value=self._draft_settings.get("appearance_mode", "Dark"))
        appearance_menu = ctk.CTkOptionMenu(appearance_frame, values=["System", "Light", "Dark"], variable=self._appearance_var)
        appearance_menu.grid(row=1, column=0, sticky="w", padx=12, pady=(0, 12))

        ctk.CTkLabel(appearance_frame, text="Тема профілю:").grid(
            row=2,
            column=0,
            sticky="w",
            padx=12,
            pady=(0, 4),
        )
        self._profile_var = tk.StringVar(value=self._draft_settings.get("theme_profile", "dark"))
        profile_menu = ctk.CTkOptionMenu(
            appearance_frame,
            values=["dark", "light"],
            variable=self._profile_var,
            command=self._on_profile_change,
        )
        profile_menu.grid(row=3, column=0, sticky="w", padx=12, pady=(0, 12))

    def _build_colors_panel(self, parent: ctk.CTkFrame) -> None:
        panel = ctk.CTkFrame(parent)
        panel.pack(fill="both", expand=True)
        self._panels["Кольори"] = panel

        self._color_vars: Dict[str, tk.StringVar] = {}

        top_actions = ctk.CTkFrame(panel)
        top_actions.pack(fill="x", padx=12, pady=(12, 6))
        top_actions.grid_columnconfigure(0, weight=1)

        copy_dark_btn = ctk.CTkButton(
            top_actions,
            text="Скопіювати dark → light",
            command=lambda: self._copy_profile_colors("dark", "light"),
        )
        copy_dark_btn.grid(row=0, column=0, sticky="w", padx=8, pady=8)
        copy_light_btn = ctk.CTkButton(
            top_actions,
            text="Скопіювати light → dark",
            command=lambda: self._copy_profile_colors("light", "dark"),
        )
        copy_light_btn.grid(row=0, column=1, sticky="w", padx=8, pady=8)
        reset_profile_btn = ctk.CTkButton(
            top_actions,
            text="Скинути поточний профіль",
            command=self._reset_current_profile,
        )
        reset_profile_btn.grid(row=0, column=2, sticky="e", padx=8, pady=8)

        scroll = ctk.CTkScrollableFrame(panel)
        scroll.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        scroll.grid_columnconfigure(1, weight=1)
        self._colors_scroll = scroll
        self._bind_scrollwheel(scroll)

        sections = [
            (
                "Основні",
                [
                    ("background", "Background"),
                    ("surface", "Surface / Panel"),
                    ("widget_fg", "Widget FG"),
                    ("border", "Border"),
                    ("text", "Text"),
                    ("accent", "Accent"),
                    ("danger", "Danger"),
                ],
            ),
            (
                "Прокрутка",
                [
                    ("scrollbar_track", "Scrollbar track"),
                    ("scrollbar_thumb", "Scrollbar thumb"),
                    ("scrollbar_thumb_hover", "Scrollbar thumb hover"),
                ],
            ),
            (
                "Заголовки / шапки панелей",
                [
                    ("header_bg", "Header BG"),
                    ("header_text", "Header text"),
                    ("header_border", "Header border"),
                ],
            ),
            (
                "Виділення",
                [
                    ("selection_bg", "Selection BG"),
                    ("selection_text", "Selection text"),
                ],
            ),
            (
                "Текст / Ввід",
                [
                    ("caret", "Caret"),
                ],
            ),
        ]
        row_index = 0
        for title, labels in sections:
            section_label = ctk.CTkLabel(scroll, text=title)
            section_label.grid(row=row_index, column=0, columnspan=6, sticky="w", padx=12, pady=(12, 2))
            row_index += 1
            for key, label in labels:
                ctk.CTkLabel(scroll, text=f"{label}:").grid(
                    row=row_index,
                    column=0,
                    sticky="w",
                    padx=(12, 6),
                    pady=(6, 2),
                )
                var = tk.StringVar()
                entry = ctk.CTkEntry(scroll, textvariable=var)
                entry.grid(row=row_index, column=1, sticky="ew", padx=6, pady=(6, 2))
                entry.bind("<KeyRelease>", lambda _event, color_key=key: self._validate_color_entry(color_key))
                entry.bind("<FocusOut>", lambda _event, color_key=key: self._validate_color_entry(color_key))
                self._entry_border_colors[key] = entry.cget("border_color")

                swatch = ctk.CTkFrame(scroll, width=28, height=24, corner_radius=4)
                swatch.grid(row=row_index, column=2, sticky="w", padx=6, pady=(6, 2))
                swatch.grid_propagate(False)
                swatch.bind("<Button-1>", lambda _event, color_key=key: self._pick_color(color_key))

                pick_btn = ctk.CTkButton(
                    scroll,
                    text="...",
                    width=40,
                    command=lambda color_key=key: self._pick_color(color_key),
                )
                pick_btn.grid(row=row_index, column=3, sticky="w", padx=6, pady=(6, 2))
                reset_btn = ctk.CTkButton(
                    scroll,
                    text="Reset",
                    width=60,
                    command=lambda color_key=key: self._reset_color(color_key),
                )
                reset_btn.grid(row=row_index, column=4, sticky="w", padx=6, pady=(6, 2))
                copy_btn = ctk.CTkButton(
                    scroll,
                    text="Copy",
                    width=60,
                    command=lambda color_key=key: self._copy_color(color_key),
                )
                copy_btn.grid(row=row_index, column=5, sticky="w", padx=6, pady=(6, 2))

                error_label = ctk.CTkLabel(scroll, text="", text_color="red")
                error_label.grid(row=row_index + 1, column=1, columnspan=5, sticky="w", padx=6, pady=(0, 4))

                self._color_vars[key] = var
                self._color_entries[key] = entry
                self._color_swatches[key] = swatch
                self._color_error_labels[key] = error_label
                row_index += 2

    def _build_fonts_panel(self, parent: ctk.CTkFrame) -> None:
        panel = ctk.CTkFrame(parent)
        panel.pack(fill="both", expand=True)
        self._panels["Шрифти"] = panel

        scroll = ctk.CTkScrollableFrame(panel)
        scroll.pack(fill="both", expand=True, padx=12, pady=12)
        self._fonts_scroll = scroll
        self._bind_scrollwheel(scroll)

        ctk.CTkLabel(scroll, text="Family:").pack(anchor="w", padx=12, pady=(12, 4))
        families = sorted(set(tkfont.families(self)))
        self._font_family_var = tk.StringVar()
        family_combo = ttk.Combobox(scroll, values=families, textvariable=self._font_family_var)
        family_combo.pack(fill="x", padx=12, pady=(0, 8))

        ctk.CTkLabel(scroll, text="Base size:").pack(anchor="w", padx=12, pady=(8, 4))
        self._font_base_var = tk.StringVar()
        base_entry = ctk.CTkEntry(scroll, textvariable=self._font_base_var)
        base_entry.pack(fill="x", padx=12, pady=(0, 8))

        ctk.CTkLabel(scroll, text="Heading size:").pack(anchor="w", padx=12, pady=(8, 4))
        self._font_heading_var = tk.StringVar()
        heading_entry = ctk.CTkEntry(scroll, textvariable=self._font_heading_var)
        heading_entry.pack(fill="x", padx=12, pady=(0, 8))

    def _select_category(self, name: str) -> None:
        for panel_name, panel in self._panels.items():
            if panel_name == name:
                panel.pack(fill="both", expand=True)
            else:
                panel.pack_forget()
        if name in {"Кольори", "Шрифти"}:
            self._refresh_profile_fields()
        if name == "Кольори" and self._colors_scroll:
            try:
                self._colors_scroll._parent_canvas.yview_moveto(0)
            except Exception:
                pass
        if name == "Шрифти" and self._fonts_scroll:
            try:
                self._fonts_scroll._parent_canvas.yview_moveto(0)
            except Exception:
                pass

    def _bind_scrollwheel(self, scroll: ctk.CTkScrollableFrame) -> None:
        try:
            canvas = scroll._parent_canvas
        except Exception:
            return

        def _on_mousewheel(event: tk.Event) -> str:
            if getattr(event, "delta", 0):
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            return "break"

        def _on_button4(_event: tk.Event) -> str:
            canvas.yview_scroll(-1, "units")
            return "break"

        def _on_button5(_event: tk.Event) -> str:
            canvas.yview_scroll(1, "units")
            return "break"

        try:
            canvas.bind("<MouseWheel>", _on_mousewheel)
            canvas.bind("<Button-4>", _on_button4)
            canvas.bind("<Button-5>", _on_button5)
        except Exception:
            pass

    def _on_category_select(self, _event: tk.Event) -> None:
        selection = self._category_list.curselection()
        if not selection:
            return
        name = self._category_list.get(selection[0])
        self._select_category(name)

    def _on_profile_change(self, _value: str) -> None:
        self._collect_profile_fields()
        self._refresh_profile_fields()

    def _refresh_profile_fields(self) -> None:
        profile = self._profile_var.get() or "dark"
        themes = self._draft_settings.setdefault("themes", {})
        theme = themes.setdefault(profile, {})
        colors = theme.setdefault("colors", {})
        fonts = theme.setdefault("fonts", {})
        default_colors = self._default_settings["themes"][profile]["colors"]

        for key, var in self._color_vars.items():
            var.set(colors.get(key, default_colors.get(key, "")))

        self._font_family_var.set(str(fonts.get("family", "")))
        self._font_base_var.set(str(fonts.get("base_size", "")))
        self._font_heading_var.set(str(fonts.get("heading_size", "")))
        self._validate_all_colors()
        self._apply_listbox_theme()

    def _collect_profile_fields(self) -> None:
        profile = self._profile_var.get() or "dark"
        themes = self._draft_settings.setdefault("themes", {})
        theme = themes.setdefault(profile, {})
        colors = theme.setdefault("colors", {})
        fonts = theme.setdefault("fonts", {})

        for key, var in self._color_vars.items():
            colors[key] = var.get().strip()

        fonts["family"] = self._font_family_var.get().strip()
        fonts["base_size"] = self._font_base_var.get().strip()
        fonts["heading_size"] = self._font_heading_var.get().strip()

    def _collect_common_fields(self) -> None:
        self._draft_settings["appearance_mode"] = self._appearance_var.get()
        self._draft_settings["theme_profile"] = self._profile_var.get()
        self._draft_settings["export_folder"] = self._export_folder_var.get().strip()

    def _apply_internal(self) -> None:
        if self._color_errors:
            return False
        try:
            self._collect_profile_fields()
            self._collect_common_fields()
            if self._on_apply:
                self._on_apply(deepcopy(self._draft_settings))
        except Exception:
            logger.exception("Не вдалося застосувати налаштування")
            messagebox.showerror("Помилка", "Не вдалося застосувати налаштування.")
            return False
        return True

    def _on_ok(self) -> None:
        if self._apply_internal():
            self.destroy()

    def _on_cancel(self) -> None:
        self.destroy()

    def _on_apply_click(self) -> None:
        self._apply_internal()

    def _reset_all(self) -> None:
        self._draft_settings = deepcopy(self._default_settings)
        self._appearance_var.set(self._draft_settings.get("appearance_mode", "Dark"))
        self._profile_var.set(self._draft_settings.get("theme_profile", "dark"))
        self._export_folder_var.set(self._draft_settings.get("export_folder", ""))
        self._refresh_profile_fields()
        self._update_action_buttons_state()

    def _choose_export_folder(self) -> None:
        path = filedialog.askdirectory(parent=self)
        if path:
            self._export_folder_var.set(path)

    def _on_cancel_event(self, _event: tk.Event) -> None:
        self._on_cancel()

    def _on_apply_shortcut(self, _event: tk.Event) -> str:
        if not self._color_errors:
            self._on_apply_click()
        return "break"

    def _on_ok_shortcut(self, _event: tk.Event) -> str:
        if not self._color_errors:
            self._on_ok()
        return "break"

    def _reset_color(self, key: str) -> None:
        profile = self._profile_var.get() or "dark"
        default_color = self._default_settings["themes"][profile]["colors"].get(key, "")
        self._color_vars[key].set(default_color)
        self._validate_color_entry(key)

    def _copy_color(self, key: str) -> None:
        value = self._color_vars[key].get().strip()
        normalized = normalize_hex_color(value)
        if not normalized:
            return
        self.clipboard_clear()
        self.clipboard_append(normalized)

    def _pick_color(self, key: str) -> None:
        current = normalize_hex_color(self._color_vars[key].get().strip())
        result = colorchooser.askcolor(parent=self, initialcolor=current or None)
        if not result or not result[1]:
            return
        normalized = normalize_hex_color(result[1])
        if not normalized:
            return
        self._color_vars[key].set(normalized)
        self._validate_color_entry(key)

    def _copy_profile_colors(self, source: str, target: str) -> None:
        themes = self._draft_settings.setdefault("themes", {})
        source_colors = themes.get(source, {}).get("colors")
        if not isinstance(source_colors, dict):
            source_colors = self._default_settings["themes"][source]["colors"]
        themes.setdefault(target, {})["colors"] = deepcopy(source_colors)
        if (self._profile_var.get() or "dark") == target:
            self._refresh_profile_fields()

    def _reset_current_profile(self) -> None:
        profile = self._profile_var.get() or "dark"
        themes = self._draft_settings.setdefault("themes", {})
        theme = themes.setdefault(profile, {})
        theme["colors"] = deepcopy(self._default_settings["themes"][profile]["colors"])
        self._refresh_profile_fields()

    def _validate_color_entry(self, key: str) -> None:
        value = self._color_vars[key].get().strip()
        normalized = normalize_hex_color(value)
        entry = self._color_entries[key]
        error_label = self._color_error_labels[key]
        if normalized:
            if value != normalized:
                self._color_vars[key].set(normalized)
            self._color_errors.discard(key)
            entry.configure(border_color=self._entry_border_colors[key])
            error_label.configure(text="")
            self._color_swatches[key].configure(fg_color=normalized)
        else:
            self._color_errors.add(key)
            entry.configure(border_color="red")
            error_label.configure(text="Очікується #RRGGBB")
        self._update_action_buttons_state()

    def _validate_all_colors(self) -> None:
        self._color_errors.clear()
        for key in self._color_vars:
            self._validate_color_entry(key)

    def _update_action_buttons_state(self) -> None:
        state = "normal" if not self._color_errors else "disabled"
        self._ok_btn.configure(state=state)
        self._apply_btn.configure(state=state)

    def _apply_listbox_theme(self) -> None:
        profile = self._profile_var.get() or "dark"
        colors = self._draft_settings.get("themes", {}).get(profile, {}).get("colors", {})
        default_colors = self._default_settings["themes"][profile]["colors"]
        if not isinstance(colors, dict):
            colors = {}
        def _color(key: str, fallback_key: str | None = None) -> str:
            return (
                normalize_hex_color(colors.get(key))
                or normalize_hex_color(default_colors.get(key))
                or (normalize_hex_color(default_colors.get(fallback_key)) if fallback_key else "")
                or ""
            )
        try:
            self._category_list.configure(
                background=_color("widget_fg"),
                foreground=_color("text"),
                selectbackground=_color("selection_bg"),
                selectforeground=_color("selection_text"),
                highlightbackground=_color("border"),
                highlightcolor=_color("border"),
            )
        except Exception:
            logger.exception("Не вдалося застосувати тему до списку категорій в SettingsDialog")
