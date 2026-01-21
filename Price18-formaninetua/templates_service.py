from __future__ import annotations

import re
from collections import Counter
from typing import Any, Callable, Iterable, Mapping

from jinja2 import Template


_NON_ALNUM_RE = re.compile(r"[^A-Za-z0-9]")


def clean_id_value(value: object) -> str:
    if value is None:
        return ""
    raw_value = str(value).strip()
    cleaned = _NON_ALNUM_RE.sub("", raw_value)
    return cleaned.upper()


def generate_export_rows(
    items: Iterable[Mapping[str, Any]],
    export_fields: Iterable[Mapping[str, str]],
    column_order: Iterable[str],
    show_info: Callable[[str], None] | None = None,
) -> list[list[str]]:
    rows: list[list[str]] = []

    for item in items:
        def spec(field_name: str) -> Any:
            return item.get(field_name)

        context = {
            "spec": spec,
            "clean_id": clean_id_value,
        }

        row: list[str] = []
        for field in export_fields:
            template_value = field.get("template", "")
            tpl = Template(template_value)
            value = tpl.render(**context)
            row.append(str(value))
        rows.append(row)

    column_list = list(column_order)
    if "Код_товару" in column_list:
        code_index = column_list.index("Код_товару")
        codes = [row[code_index] for row in rows if len(row) > code_index]
        duplicates = [code for code, count in Counter(codes).items() if code and count > 1]
        if duplicates:
            message = (
                "Warning: duplicate values found in 'Код_товару' after cleaning: "
                + ", ".join(sorted(duplicates))
            )
            if show_info is not None:
                show_info(message)
    return rows
