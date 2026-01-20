"""Utilities for importing and exporting model specifications."""
from __future__ import annotations

from typing import List, Sequence, Tuple

_HEADERS = {
    "key",
    "назва",
    "назва параметра",
    "характеристика",
    "parameter",
    "name",
}

_SEPARATORS = ("\t", ";", ",", ":", "=")


def parse_specs_payload(raw: str) -> List[Tuple[str, str]]:
    """Parse plain text or CSV-like payload into key/value pairs.

    The parser is intentionally forgiving – it splits on a first available
    separator and trims whitespace/quotes. Header rows with well-known column
    names are ignored.
    """

    if not raw:
        return []
    text = raw.replace("\r\n", "\n").replace("\r", "\n")
    lines = text.split("\n")
    pairs: List[Tuple[str, str]] = []
    first_data_row = True
    for line in lines:
        candidate = line.strip()
        if candidate.startswith("\ufeff"):
            candidate = candidate.lstrip("\ufeff")
        if not candidate:
            continue
        if candidate.startswith("#"):
            # allow comments in pasted snippets
            continue
        key = candidate
        value = ""
        for separator in _SEPARATORS:
            if separator in candidate:
                key_part, value_part = candidate.split(separator, 1)
                key = key_part
                value = value_part
                break
        key = key.strip().strip('"').strip("'").rstrip(":;,=")
        if key.startswith("\ufeff"):
            key = key.lstrip("\ufeff")
        if not key:
            continue
        lower_key = key.lower()
        if first_data_row and lower_key in _HEADERS:
            # skip header line
            continue
        value = value.strip().strip('"').strip("'")
        if value.startswith("\ufeff"):
            value = value.lstrip("\ufeff")
        pairs.append((key, value))
        first_data_row = False
    return pairs


def format_specs_for_clipboard(specs: Sequence[Tuple[str, str]]) -> str:
    """Format specs into a tab-delimited string for clipboard."""

    lines: List[str] = []
    for key, value in specs:
        safe_key = (key or "").strip()
        safe_value = (value or "").strip()
        lines.append(f"{safe_key}\t{safe_value}")
    return "\n".join(lines)
