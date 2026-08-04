"""Interpretador tolerante para pares referência/valor em texto livre."""
from __future__ import annotations

import re
from typing import Any


SEPARATOR = re.compile(r"\s*(?:->|=|:|;|,|\||-|\t|\s{2,})\s*")
NUMBER = re.compile(r"^-?\d+(?:[.,]\d+)?$")


def _number(value: str) -> int | float:
    normalized = value.replace(".", "").replace(",", ".") if "," in value else value
    number = float(normalized)
    return int(number) if number.is_integer() else number


def parse_entries(text: str) -> tuple[list[dict[str, Any]], list[str]]:
    """Accept forms like `S21 - 6`, `S21:6` and `Produto A  15`."""
    rows: list[dict[str, Any]] = []
    warnings: list[str] = []
    seen: dict[str, int] = {}
    for line_number, raw_line in enumerate(text.splitlines(), start=1):
        line = raw_line.strip()
        if not line:
            continue
        pieces = [piece.strip() for piece in SEPARATOR.split(line) if piece.strip()]
        if len(pieces) < 2:
            # Final whitespace fallback supports "S21 6" but keeps multiword names intact when possible.
            match = re.match(r"^(.+?)\s+(-?\d+(?:[.,]\d+)?)$", line)
            pieces = [match.group(1), match.group(2)] if match else []
        if len(pieces) < 2 or not NUMBER.match(pieces[-1]):
            warnings.append(f"Linha {line_number} ignorada: '{raw_line}'.")
            continue
        reference = " ".join(pieces[:-1]).strip()
        if not reference:
            warnings.append(f"Linha {line_number} ignorada: referência vazia.")
            continue
        if reference.upper() in {"CODIGO", "CÓDIGO", "REFERENCIA", "REFERÊNCIA"}:
            continue
        row = {"Referência": reference, "Valor": _number(pieces[-1])}
        key = reference.casefold()
        if key in seen:
            rows[seen[key]] = row
            warnings.append(f"Referência repetida '{reference}': foi mantido o último valor.")
        else:
            seen[key] = len(rows)
            rows.append(row)
    return rows, warnings
