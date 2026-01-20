"""Utilities for exporting and importing application data via Excel workbooks."""

from __future__ import annotations

import json
import os
from datetime import datetime
from typing import Dict, Iterable, List, Optional, Tuple

import importlib

from database import export_catalog_dump, import_catalog_dump
from templates_service import (
    CSV_JSON_FALLBACK_NOTE,
    EXCEL_EXPORT_BLOCKED_MESSAGE,
    OPENPYXL_AVAILABLE,
    OPENPYXL_IMPORT_ERROR_DETAIL,
    OPENPYXL_INSTALL_HINT,
    Workbook,
)


class DataTransferError(RuntimeError):
    """Custom exception raised for import/export failures."""

    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code
        self.message = message

    def __str__(self) -> str:  # pragma: no cover - mirrors base behaviour
        return self.message


DATA_ERR_NO_OPENPYXL = "NO_OPENPYXL"
DATA_ERR_INVALID_PAYLOAD = "INVALID_PAYLOAD"
DATA_ERR_IO = "IO_ERROR"
DATA_ERR_SAVE_FAILED = "SAVE_FAILED"
DATA_ERR_LOAD_FAILED = "LOAD_FAILED"
DATA_ERR_VALIDATION = "VALIDATION_FAILED"


def _require_openpyxl() -> None:
    if not OPENPYXL_AVAILABLE or EXCEL_EXPORT_BLOCKED_MESSAGE:
        message = EXCEL_EXPORT_BLOCKED_MESSAGE or "Експорт у Excel недоступний."
        detail = OPENPYXL_IMPORT_ERROR_DETAIL
        if detail and detail not in message:
            message = f"{message} (деталі: {detail})"
        message = f"{message}\n{OPENPYXL_INSTALL_HINT}{CSV_JSON_FALLBACK_NOTE}"
        raise DataTransferError(DATA_ERR_NO_OPENPYXL, message)
    if Workbook is None:
        raise DataTransferError(
            DATA_ERR_NO_OPENPYXL,
            "Експорт у формат Excel недоступний у цій конфігурації.",
        )


def _load_workbook_file(path: str):
    module = importlib.import_module("openpyxl")
    return module.load_workbook(path)


def _normalise_languages(raw: Iterable[str | None]) -> List[str]:
    seen = set()
    result: List[str] = []
    for entry in raw:
        if not isinstance(entry, str):
            continue
        code = entry.strip()
        if not code or code in seen:
            continue
        result.append(code)
        seen.add(code)
    return result


def export_all_data_to_excel(
    filename: str,
    templates: Dict[str, object],
    title_tags: Dict[str, object],
    export_fields: List[Dict[str, object]],
) -> str:
    """Export catalog, templates, parameters and export fields into an Excel file."""

    _require_openpyxl()

    if not filename:
        raise DataTransferError(DATA_ERR_IO, "Не вказано шлях для збереження файлу.")

    catalog = export_catalog_dump()

    workbook = Workbook()
    sheets = {}

    sheet = workbook.active
    sheet.title = "Категорії"
    sheet.append(["ID", "Назва", "Створено"])
    for entry in catalog.get("categories", []):
        sheet.append(
            [
                entry.get("id"),
                entry.get("name"),
                entry.get("created_at"),
            ]
        )
    sheets["Категорії"] = sheet

    sheet = workbook.create_sheet("Бренди")
    sheet.append(["ID", "Категорія ID", "Назва", "Створено"])
    for entry in catalog.get("brands", []):
        sheet.append(
            [
                entry.get("id"),
                entry.get("category_id"),
                entry.get("name"),
                entry.get("created_at"),
            ]
        )
    sheets["Бренди"] = sheet

    sheet = workbook.create_sheet("Моделі")
    sheet.append(["ID", "Бренд ID", "Назва", "Створено"])
    for entry in catalog.get("models", []):
        sheet.append(
            [
                entry.get("id"),
                entry.get("brand_id"),
                entry.get("name"),
                entry.get("created_at"),
            ]
        )
    sheets["Моделі"] = sheet

    sheet = workbook.create_sheet("Характеристики")
    sheet.append(["ID", "Модель ID", "Ключ", "Значення"])
    for entry in catalog.get("specs", []):
        sheet.append(
            [
                entry.get("id"),
                entry.get("model_id"),
                entry.get("key"),
                entry.get("value"),
            ]
        )
    sheets["Характеристики"] = sheet

    sheet = workbook.create_sheet("Параметри")
    sheet.append(["Тип", "Код", "Назва", "Увімкнено"])
    for entry in templates.get("template_languages", []) if isinstance(templates, dict) else []:
        if not isinstance(entry, dict):
            continue
        code = (entry.get("code") or "").strip()
        label = (entry.get("label") or "").strip() or code
        if not code:
            continue
        sheet.append(["language", code, label, ""])
    for entry in templates.get("film_types", []) if isinstance(templates, dict) else []:
        if not isinstance(entry, dict):
            continue
        name = (entry.get("name") or "").strip()
        if not name:
            continue
        enabled = "1" if entry.get("enabled") else "0"
        sheet.append(["film_type", "", name, enabled])
    sheets["Параметри"] = sheet

    sheet = workbook.create_sheet("Шаблони")
    sheet.append([
        "Група",
        "Сценарій",
        "Поле",
        "Категорія",
        "Тип плівки",
        "Мова",
        "Значення",
    ])

    def as_text(value: object) -> str:
        if value is None:
            return ""
        if isinstance(value, str):
            return value
        return str(value)

    templates_dict = templates if isinstance(templates, dict) else {}
    title_template_value = as_text(templates_dict.get("title_template"))
    tags_template_value = as_text(templates_dict.get("tags_template"))
    sheet.append([
        "templates",
        "global",
        "title_template",
        "",
        "",
        "",
        title_template_value,
    ])
    sheet.append([
        "templates",
        "global",
        "tags_template",
        "",
        "",
        "",
        tags_template_value,
    ])

    descriptions = templates_dict.get("descriptions")
    if isinstance(descriptions, dict):
        for category, films in descriptions.items():
            if not isinstance(films, dict):
                continue
            category_name = as_text(category)
            for film_type, payload in films.items():
                if not isinstance(payload, dict):
                    continue
                film_name = as_text(film_type)
                default_text = as_text(payload.get("default"))
                sheet.append([
                    "templates",
                    "description_default",
                    "description",
                    category_name,
                    film_name,
                    "",
                    default_text,
                ])
                languages = payload.get("languages")
                if isinstance(languages, dict):
                    for code, text in languages.items():
                        code_value = as_text(code)
                        if not code_value:
                            continue
                        sheet.append([
                            "templates",
                            "description_language",
                            "description",
                            category_name,
                            film_name,
                            code_value,
                            as_text(text),
                        ])

    def append_row(
        group: str,
        scenario: str,
        field: str,
        category: str,
        film: str,
        language: str,
        value: object,
    ) -> None:
        sheet.append([
            group,
            scenario,
            field,
            category,
            film,
            language,
            as_text(value),
        ])

    def append_title_tags_entry(
        scenario: str,
        field: str,
        category: str,
        film: str,
        language: str,
        value: object,
    ) -> None:
        append_row("title_tags", scenario, field, category, film, language, value)

    def iter_template_entry(entry: object):
        default_value = None
        languages_dict: Dict[str, object] = {}
        if isinstance(entry, dict):
            default_value = entry.get("default")
            languages_raw = entry.get("languages")
            if isinstance(languages_raw, dict):
                for code, text in languages_raw.items():
                    code_str = as_text(code).strip()
                    if not code_str:
                        continue
                    languages_dict[code_str] = text
        elif entry is not None:
            default_value = entry
        return default_value, languages_dict

    title_tags_dict = title_tags if isinstance(title_tags, dict) else {}
    default_block = title_tags_dict.get("default")
    if isinstance(default_block, dict):
        for field_name in ("title_template", "tags_template"):
            default_value, languages_map = iter_template_entry(default_block.get(field_name))
            append_title_tags_entry("default", field_name, "", "", "", default_value)
            for code, text in languages_map.items():
                code_value = as_text(code)
                if not code_value:
                    continue
                append_title_tags_entry(
                    "default_language",
                    field_name,
                    "",
                    "",
                    code_value,
                    text,
                )

    by_film_block = title_tags_dict.get("by_film")
    if isinstance(by_film_block, dict):
        for film_type, payload in by_film_block.items():
            if not isinstance(payload, dict):
                continue
            film_name = as_text(film_type)
            for field_name in ("title_template", "tags_template"):
                default_value, languages_map = iter_template_entry(payload.get(field_name))
                append_title_tags_entry("film", field_name, "", film_name, "", default_value)
                for code, text in languages_map.items():
                    code_value = as_text(code)
                    if not code_value:
                        continue
                    append_title_tags_entry(
                        "film_language",
                        field_name,
                        "",
                        film_name,
                        code_value,
                        text,
                    )

    by_category_block = title_tags_dict.get("by_category")
    if isinstance(by_category_block, dict):
        for category, payload in by_category_block.items():
            if not isinstance(payload, dict):
                continue
            category_name = as_text(category)
            category_default = payload.get("default")
            if isinstance(category_default, dict):
                for field_name in ("title_template", "tags_template"):
                    default_value, languages_map = iter_template_entry(
                        category_default.get(field_name)
                    )
                    append_title_tags_entry(
                        "category_default",
                        field_name,
                        category_name,
                        "",
                        "",
                        default_value,
                    )
                    for code, text in languages_map.items():
                        code_value = as_text(code)
                        if not code_value:
                            continue
                        append_title_tags_entry(
                            "category_default_language",
                            field_name,
                            category_name,
                            "",
                            code_value,
                            text,
                        )
            category_by_film = payload.get("by_film")
            if isinstance(category_by_film, dict):
                for film_type, film_payload in category_by_film.items():
                    if not isinstance(film_payload, dict):
                        continue
                    film_name = as_text(film_type)
                    for field_name in ("title_template", "tags_template"):
                        default_value, languages_map = iter_template_entry(
                            film_payload.get(field_name)
                        )
                        append_title_tags_entry(
                            "category_film",
                            field_name,
                            category_name,
                            film_name,
                            "",
                            default_value,
                        )
                        for code, text in languages_map.items():
                            code_value = as_text(code)
                            if not code_value:
                                continue
                            append_title_tags_entry(
                                "category_film_language",
                                field_name,
                                category_name,
                                film_name,
                                code_value,
                                text,
                            )

    known_template_keys = {
        "title_template",
        "tags_template",
        "descriptions",
        "template_languages",
        "film_types",
    }
    for key, value in templates_dict.items():
        if key in known_template_keys:
            continue
        append_row(
            "templates",
            "raw_json",
            key,
            "",
            "",
            "",
            json.dumps(value, ensure_ascii=False, indent=2),
        )

    known_title_tag_keys = {"default", "by_film", "by_category"}
    for key, value in title_tags_dict.items():
        if key in known_title_tag_keys:
            continue
        append_row(
            "title_tags",
            "raw_json",
            key,
            "",
            "",
            "",
            json.dumps(value, ensure_ascii=False, indent=2),
        )

    sheets["Шаблони"] = sheet

    sheet = workbook.create_sheet("Експортні поля")
    sheet.append(["Позиція", "Поле", "Шаблон", "Увімкнено", "Мови"])
    for idx, entry in enumerate(export_fields, start=1):
        if not isinstance(entry, dict):
            continue
        field_name = (entry.get("field") or "").strip()
        template_text = entry.get("template", "")
        enabled = "1" if entry.get("enabled") else "0"
        languages_raw = entry.get("languages")
        if isinstance(languages_raw, str):
            languages_list = [languages_raw.strip()]
        elif isinstance(languages_raw, Iterable):
            languages_list = [
                str(code).strip()
                for code in languages_raw
                if isinstance(code, str) and code.strip()
            ]
        else:
            languages_list = []
        languages = ", ".join(_normalise_languages(languages_list))
        sheet.append([idx, field_name, template_text, enabled, languages])
    sheets["Експортні поля"] = sheet

    try:
        directory = os.path.dirname(filename)
        if directory and not os.path.exists(directory):
            os.makedirs(directory, exist_ok=True)
        workbook.save(filename)
    except PermissionError as exc:
        raise DataTransferError(
            DATA_ERR_IO,
            "Не вдалося зберегти Excel-файл: доступ заборонено. Закрийте файл, якщо він відкритий, та спробуйте знову.",
        ) from exc
    except OSError as exc:
        raise DataTransferError(DATA_ERR_IO, f"Не вдалося зберегти файл: {exc}") from exc
    finally:
        try:
            workbook.close()
        except Exception:  # pragma: no cover - best effort cleanup
            pass

    return filename


def _parse_parameters_sheet(sheet) -> Tuple[List[Dict[str, object]], List[Dict[str, object]]]:
    languages: List[Dict[str, object]] = []
    film_types: List[Dict[str, object]] = []
    if sheet is None:
        return languages, film_types

    rows = getattr(sheet, "iter_rows", None)
    if rows is None:
        return languages, film_types

    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        raw_type = row[0]
        entry_type = str(raw_type).strip().lower() if isinstance(raw_type, str) else ""
        if not entry_type:
            continue
        if entry_type == "language":
            code = str(row[1]).strip() if isinstance(row[1], str) else ""
            label = str(row[2]).strip() if isinstance(row[2], str) else ""
            if not code:
                continue
            languages.append({"code": code, "label": label or code})
        elif entry_type == "film_type":
            name = ""
            if isinstance(row[2], str):
                name = row[2].strip()
            if not name and isinstance(row[1], str):
                name = row[1].strip()
            if not name:
                continue
            enabled_cell = row[3] if len(row) > 3 else None
            enabled_value = False
            if isinstance(enabled_cell, str):
                enabled_value = enabled_cell.strip().lower() in {"1", "true", "yes", "y", "так"}
            elif isinstance(enabled_cell, (int, float)):
                enabled_value = bool(enabled_cell)
            elif isinstance(enabled_cell, bool):
                enabled_value = enabled_cell
            film_types.append({"name": name, "enabled": enabled_value})

    return languages, film_types


def _parse_templates_sheet(sheet) -> Tuple[Optional[Dict[str, object]], Optional[Dict[str, object]]]:
    templates_data: Optional[Dict[str, object]] = None
    title_tags_data: Optional[Dict[str, object]] = None
    if sheet is None:
        return templates_data, title_tags_data

    rows_iter = getattr(sheet, "iter_rows", None)
    if rows_iter is None:
        return templates_data, title_tags_data

    rows = [tuple(row) for row in sheet.iter_rows(min_row=2, values_only=True)]

    def is_new_format(row: Tuple[object, ...]) -> bool:
        if not row:
            return False
        if len(row) >= 4 and isinstance(row[0], str):
            return True
        return False

    if not rows:
        return templates_data, title_tags_data

    if any(is_new_format(row) for row in rows):
        return _parse_templates_sheet_new(rows)

    # Fallback to legacy JSON-based format for backwards compatibility
    for row in rows:
        if not row:
            continue
        key = row[0]
        if not isinstance(key, str):
            continue
        payload = row[1] if len(row) > 1 else None
        if not isinstance(payload, str):
            continue
        try:
            data = json.loads(payload)
        except json.JSONDecodeError as exc:
            raise DataTransferError(DATA_ERR_VALIDATION, f"Некоректний JSON у блоці '{key}': {exc}")
        if key == "templates":
            if not isinstance(data, dict):
                raise DataTransferError(DATA_ERR_VALIDATION, "Розділ 'templates' повинен містити об'єкт JSON.")
            templates_data = data
        elif key == "title_tags_templates":
            if not isinstance(data, dict):
                raise DataTransferError(
                    DATA_ERR_VALIDATION,
                    "Розділ 'title_tags_templates' повинен містити об'єкт JSON.",
                )
            title_tags_data = data

    return templates_data, title_tags_data


def _parse_templates_sheet_new(
    rows: List[Tuple[object, ...]]
) -> Tuple[Optional[Dict[str, object]], Optional[Dict[str, object]]]:
    templates_result: Dict[str, object] = {}
    descriptions: Dict[str, Dict[str, Dict[str, object]]] = {}

    title_tags_default: Dict[str, Dict[str, object]] = {}
    title_tags_by_film: Dict[str, Dict[str, Dict[str, object]]] = {}
    title_tags_by_category: Dict[str, Dict[str, object]] = {}
    extra_title_tags: Dict[str, object] = {}

    def cell_to_key(value: object, *, lower: bool = False) -> str:
        if value is None:
            return ""
        text = str(value).strip()
        return text.lower() if lower else text

    def value_text(value: object) -> str:
        if value is None:
            return ""
        if isinstance(value, str):
            return value
        return str(value)

    def ensure_description(category: str, film: str) -> Dict[str, object]:
        category_map = descriptions.setdefault(category, {})
        film_entry = category_map.setdefault(film, {})
        languages = film_entry.get("languages")
        if not isinstance(languages, dict):
            languages = {}
        film_entry["languages"] = languages
        return film_entry

    def ensure_template_entry(container: Dict[str, Dict[str, object]], field: str) -> Dict[str, object]:
        entry = container.get(field)
        if not isinstance(entry, dict):
            entry = {}
        languages = entry.get("languages")
        if not isinstance(languages, dict):
            languages = {}
        entry["languages"] = languages
        container[field] = entry
        return entry

    for row in rows:
        if not row:
            continue
        group = cell_to_key(row[0], lower=True)
        if not group:
            continue
        scenario = cell_to_key(row[1] if len(row) > 1 else "", lower=True)
        field = cell_to_key(row[2] if len(row) > 2 else "")
        category = cell_to_key(row[3] if len(row) > 3 else "")
        film = cell_to_key(row[4] if len(row) > 4 else "")
        language = cell_to_key(row[5] if len(row) > 5 else "")
        raw_value = row[6] if len(row) > 6 else ""
        text_value = value_text(raw_value)

        if group == "templates":
            if scenario == "global":
                if field:
                    templates_result[field] = text_value
            elif scenario == "description_default":
                if category and film:
                    entry = ensure_description(category, film)
                    entry["default"] = text_value
            elif scenario == "description_language":
                if category and film and language:
                    entry = ensure_description(category, film)
                    languages_map = entry.setdefault("languages", {})
                    languages_map[language] = text_value
            elif scenario == "raw_json":
                if not field:
                    continue
                payload = text_value.strip()
                if not payload:
                    templates_result[field] = None
                else:
                    try:
                        templates_result[field] = json.loads(payload)
                    except json.JSONDecodeError as exc:
                        raise DataTransferError(
                            DATA_ERR_VALIDATION,
                            f"Некоректний JSON у полі '{field}' (група 'templates'): {exc}",
                        ) from exc
            else:
                if field:
                    templates_result[field] = text_value
        elif group == "title_tags":
            if scenario == "default":
                if field:
                    entry = ensure_template_entry(title_tags_default, field)
                    entry["default"] = text_value
            elif scenario == "default_language":
                if field and language:
                    entry = ensure_template_entry(title_tags_default, field)
                    entry.setdefault("languages", {})[language] = text_value
            elif scenario == "film":
                if field and film:
                    film_entry = ensure_template_entry(
                        title_tags_by_film.setdefault(film, {}), field
                    )
                    film_entry["default"] = text_value
            elif scenario == "film_language":
                if field and film and language:
                    film_entry = ensure_template_entry(
                        title_tags_by_film.setdefault(film, {}), field
                    )
                    film_entry.setdefault("languages", {})[language] = text_value
            elif scenario == "category_default":
                if field and category:
                    container = ensure_template_entry(
                        title_tags_by_category.setdefault(category, {}).setdefault(
                            "default", {}
                        ),
                        field,
                    )
                    container["default"] = text_value
            elif scenario == "category_default_language":
                if field and category and language:
                    container = ensure_template_entry(
                        title_tags_by_category.setdefault(category, {}).setdefault(
                            "default", {}
                        ),
                        field,
                    )
                    container.setdefault("languages", {})[language] = text_value
            elif scenario == "category_film":
                if field and category and film:
                    container = ensure_template_entry(
                        title_tags_by_category.setdefault(category, {})
                        .setdefault("by_film", {})
                        .setdefault(film, {}),
                        field,
                    )
                    container["default"] = text_value
            elif scenario == "category_film_language":
                if field and category and film and language:
                    container = ensure_template_entry(
                        title_tags_by_category.setdefault(category, {})
                        .setdefault("by_film", {})
                        .setdefault(film, {}),
                        field,
                    )
                    container.setdefault("languages", {})[language] = text_value
            elif scenario == "raw_json":
                if not field:
                    continue
                payload = text_value.strip()
                if not payload:
                    extra_title_tags[field] = None
                else:
                    try:
                        extra_title_tags[field] = json.loads(payload)
                    except json.JSONDecodeError as exc:
                        raise DataTransferError(
                            DATA_ERR_VALIDATION,
                            f"Некоректний JSON у полі '{field}' (група 'title_tags'): {exc}",
                        ) from exc

    if descriptions:
        templates_result["descriptions"] = descriptions

    def clean_template_holder(holder: Dict[str, Dict[str, object]]) -> Dict[str, Dict[str, object]]:
        cleaned: Dict[str, Dict[str, object]] = {}
        for field, data in holder.items():
            if not isinstance(data, dict):
                continue
            entry: Dict[str, object] = {}
            if "default" in data:
                entry["default"] = data.get("default", "")
            languages_map = data.get("languages")
            if isinstance(languages_map, dict):
                entry["languages"] = dict(languages_map)
            elif "languages" in data:
                entry["languages"] = {}
            if entry:
                cleaned[field] = entry
        return cleaned

    title_tags_result: Dict[str, object] = {}
    if title_tags_default:
        cleaned_default = clean_template_holder(title_tags_default)
        if cleaned_default:
            title_tags_result["default"] = cleaned_default
    if title_tags_by_film:
        cleaned_by_film: Dict[str, Dict[str, object]] = {}
        for film, holder in title_tags_by_film.items():
            cleaned_holder = clean_template_holder(holder)
            if cleaned_holder:
                cleaned_by_film[film] = cleaned_holder
        if cleaned_by_film:
            title_tags_result["by_film"] = cleaned_by_film
    if title_tags_by_category:
        cleaned_by_category: Dict[str, Dict[str, object]] = {}
        for category, payload in title_tags_by_category.items():
            cleaned_category: Dict[str, object] = {}
            default_block = payload.get("default")
            if isinstance(default_block, dict):
                cleaned_default = clean_template_holder(default_block)
                if cleaned_default:
                    cleaned_category["default"] = cleaned_default
            by_film_block = payload.get("by_film")
            if isinstance(by_film_block, dict):
                cleaned_nested: Dict[str, Dict[str, object]] = {}
                for film, holder in by_film_block.items():
                    cleaned_holder = clean_template_holder(holder)
                    if cleaned_holder:
                        cleaned_nested[film] = cleaned_holder
                if cleaned_nested:
                    cleaned_category["by_film"] = cleaned_nested
            if cleaned_category:
                cleaned_by_category[category] = cleaned_category
        if cleaned_by_category:
            title_tags_result["by_category"] = cleaned_by_category

    if extra_title_tags:
        for key, value in extra_title_tags.items():
            title_tags_result[key] = value

    templates_payload = templates_result if templates_result else None
    title_tags_payload = title_tags_result if title_tags_result else None
    return templates_payload, title_tags_payload


def _parse_export_fields_sheet(sheet) -> List[Dict[str, object]]:
    if sheet is None:
        return []
    rows = getattr(sheet, "iter_rows", None)
    if rows is None:
        return []

    fields: List[Tuple[int, Dict[str, object]]] = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        position_raw = row[0]
        try:
            position = int(position_raw)
        except (TypeError, ValueError):
            position = len(fields) + 1
        field_name = (row[1] or "").strip() if isinstance(row[1], str) else ""
        template_text = row[2] if isinstance(row[2], str) else ""
        enabled_cell = row[3] if len(row) > 3 else None
        languages_cell = row[4] if len(row) > 4 else ""

        enabled = False
        if isinstance(enabled_cell, str):
            enabled = enabled_cell.strip().lower() in {"1", "true", "yes", "y", "так"}
        elif isinstance(enabled_cell, (int, float)):
            enabled = bool(enabled_cell)
        elif isinstance(enabled_cell, bool):
            enabled = enabled_cell

        if isinstance(languages_cell, str):
            raw_codes = [part.strip() for part in languages_cell.split(",")]
        elif isinstance(languages_cell, Iterable):
            raw_codes = [str(part).strip() for part in languages_cell]
        else:
            raw_codes = []
        languages = _normalise_languages(raw_codes)

        fields.append(
            (
                position,
                {
                    "field": field_name,
                    "template": template_text,
                    "enabled": enabled,
                    "languages": languages,
                },
            )
        )

    fields.sort(key=lambda item: item[0])
    return [payload for _pos, payload in fields]


def import_all_data_from_excel(filename: str) -> Dict[str, object]:
    """Import catalog, templates and export settings from an Excel workbook."""

    _require_openpyxl()

    if not filename:
        raise DataTransferError(DATA_ERR_LOAD_FAILED, "Не вказано файл для імпорту.")
    if not os.path.exists(filename):
        raise DataTransferError(DATA_ERR_LOAD_FAILED, "Файл не знайдено.")

    try:
        workbook = _load_workbook_file(filename)
    except FileNotFoundError as exc:
        raise DataTransferError(DATA_ERR_LOAD_FAILED, "Файл не знайдено.") from exc
    except OSError as exc:
        raise DataTransferError(DATA_ERR_LOAD_FAILED, f"Не вдалося відкрити файл: {exc}") from exc

    try:
        sheets = {name: workbook[name] for name in getattr(workbook, "sheetnames", [])}

        templates_sheet = sheets.get("Шаблони")
        parameters_sheet = sheets.get("Параметри")
        export_sheet = sheets.get("Експортні поля")

        catalog_payload = {
            "categories": [],
            "brands": [],
            "models": [],
            "specs": [],
        }

        if "Категорії" in sheets:
            sheet = sheets["Категорії"]
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if not row:
                    continue
                catalog_payload["categories"].append(
                    {"id": row[0], "name": row[1], "created_at": row[2]}
                )

        if "Бренди" in sheets:
            sheet = sheets["Бренди"]
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if not row:
                    continue
                catalog_payload["brands"].append(
                    {
                        "id": row[0],
                        "category_id": row[1],
                        "name": row[2],
                        "created_at": row[3] if len(row) > 3 else None,
                    }
                )

        if "Моделі" in sheets:
            sheet = sheets["Моделі"]
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if not row:
                    continue
                catalog_payload["models"].append(
                    {
                        "id": row[0],
                        "brand_id": row[1],
                        "name": row[2],
                        "created_at": row[3] if len(row) > 3 else None,
                    }
                )

        if "Характеристики" in sheets:
            sheet = sheets["Характеристики"]
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if not row:
                    continue
                catalog_payload["specs"].append(
                    {
                        "id": row[0],
                        "model_id": row[1],
                        "key": row[2],
                        "value": row[3] if len(row) > 3 else None,
                    }
                )

        templates_data, title_tags_data = _parse_templates_sheet(templates_sheet)
        languages, film_types = _parse_parameters_sheet(parameters_sheet)
        export_fields = _parse_export_fields_sheet(export_sheet)

        if templates_data is None:
            templates_data = {}
        if languages:
            templates_data["template_languages"] = languages
        if film_types:
            templates_data["film_types"] = film_types
        if title_tags_data is None:
            title_tags_data = {}

        import_catalog_dump(catalog_payload)

        from templates_service import (
            load_export_fields,
            load_templates,
            load_title_tags_templates,
            save_export_fields,
            save_templates,
            save_title_tags_templates,
        )

        try:
            save_templates(templates_data)
            save_title_tags_templates(title_tags_data)
            if export_fields:
                save_export_fields(export_fields)
        except OSError as exc:
            raise DataTransferError(DATA_ERR_SAVE_FAILED, f"Не вдалося зберегти дані: {exc}") from exc

        refreshed_templates = load_templates()
        refreshed_title_tags = load_title_tags_templates(refreshed_templates)
        refreshed_export_fields = load_export_fields()

        return {
            "templates": refreshed_templates,
            "title_tags_templates": refreshed_title_tags,
            "export_fields": refreshed_export_fields,
        }
    finally:
        try:
            workbook.close()
        except Exception:  # pragma: no cover - defensive close
            pass
