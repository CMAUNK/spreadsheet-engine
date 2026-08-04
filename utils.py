"""Pequenas utilidades compartilhadas."""
from __future__ import annotations

import re
from typing import Any


def normalize_reference(value: Any) -> str:
    return " ".join(str(value or "").strip().casefold().split())


def column_label(index: int) -> str:
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def safe_output_name(name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", name.strip())
    return cleaned.strip("._") or "smartsheet"
