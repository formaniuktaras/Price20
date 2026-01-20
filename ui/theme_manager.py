"""Apply appearance mode, fonts, and colors to key UI widgets."""
from __future__ import annotations

from typing import Any, Dict, List, Tuple
import logging

import tkinter as tk
import tkinter.font as tkfont

import customtkinter as ctk

logger = logging.getLogger(__name__)


class ThemeManager:
    def __init__(self, root: ctk.CTk) -> None:
        self.root = root
        self.widgets: List[Tuple[tk.Widget, str]] = []
        self.colors: Dict[str, str] = {}
        self.fonts: Dict[str, Any] = {}
        self.base_font: ctk.CTkFont | None = None
        self.heading_font: ctk.CTkFont | None = None

    def register(self, widget: tk.Widget, role: str) -> None:
        self.widgets.append((widget, role))

    def _resolve_font_family(self, family: str) -> str:
        try:
            families = set(tkfont.families(self.root))
        except Exception:
            families = set()
        if family in families:
            return family
        return next(iter(families), family)

    def apply(self, settings: Dict[str, Any], *, apply_widgets: bool = True) -> None:
        appearance = settings.get("appearance_mode", "Dark")
        ctk.set_appearance_mode(appearance)

        profile = settings.get("theme_profile", "dark")
        theme = settings.get("themes", {}).get(profile, {})
        self.colors = theme.get("colors", {})
        self.fonts = theme.get("fonts", {})

        family = self._resolve_font_family(str(self.fonts.get("family", "Segoe UI")))
        base_size = int(self.fonts.get("base_size", 12))
        heading_size = int(self.fonts.get("heading_size", 14))
        self.base_font = ctk.CTkFont(family=family, size=base_size)
        self.heading_font = ctk.CTkFont(family=family, size=heading_size, weight="bold")

        if not apply_widgets:
            return

        for widget, role in self.widgets:
            self._apply_widget(widget, role)

    def _apply_widget(self, widget: tk.Widget, role: str) -> None:
        colors = self.colors
        base_font = self.base_font
        heading_font = self.heading_font

        def safe_configure(**kwargs: Any) -> None:
            try:
                widget.configure(**kwargs)
            except Exception:
                logger.exception("Не вдалося застосувати тему до %s (%s)", widget, role)

        if role == "background":
            safe_configure(fg_color=colors.get("background"))
        elif role == "surface":
            safe_configure(fg_color=colors.get("surface"))
        elif role == "widget":
            safe_configure(
                fg_color=colors.get("widget_fg"),
                text_color=colors.get("text"),
                border_color=colors.get("border"),
            )
        elif role == "label":
            safe_configure(text_color=colors.get("text"))
        elif role == "menu_button":
            safe_configure(
                fg_color="transparent",
                hover_color=colors.get("widget_fg"),
                text_color=colors.get("text"),
                border_color=colors.get("border"),
            )
        elif role == "accent_button":
            safe_configure(
                fg_color=colors.get("accent"),
                hover_color=colors.get("accent"),
                text_color="#ffffff",
            )
        elif role == "danger_button":
            safe_configure(
                fg_color=colors.get("danger"),
                hover_color=colors.get("danger"),
                text_color="#ffffff",
            )
        elif role == "tabview":
            safe_configure(
                fg_color=colors.get("background"),
                segmented_button_fg_color=colors.get("surface"),
                segmented_button_selected_color=colors.get("accent"),
            )

        if base_font and role in {"widget", "label", "menu_button", "accent_button", "danger_button"}:
            safe_configure(font=base_font)
        if heading_font and role == "heading_label":
            safe_configure(font=heading_font)
