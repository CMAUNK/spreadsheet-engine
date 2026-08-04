"""Persistência local e simples de configurações por modelo de arquivo."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any


CONFIG_PATH = Path(__file__).parent / "modelos" / "configuracoes.json"
UNITS_PATH = Path(__file__).parent / "modelos" / "unidades.json"


def _read_all() -> dict[str, Any]:
    if not CONFIG_PATH.exists():
        return {}
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def load_template(fingerprint: str) -> dict[str, Any]:
    return _read_all().get(fingerprint, {})


def save_template(fingerprint: str, configuration: dict[str, Any]) -> None:
    CONFIG_PATH.parent.mkdir(exist_ok=True)
    data = _read_all()
    data[fingerprint] = configuration
    CONFIG_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def load_units(model_key: str) -> list[dict[str, Any]]:
    if not UNITS_PATH.exists():
        return []
    try:
        return json.loads(UNITS_PATH.read_text(encoding="utf-8")).get(model_key, [])
    except (OSError, json.JSONDecodeError):
        return []


def save_units(model_key: str, units: list[dict[str, Any]]) -> None:
    UNITS_PATH.parent.mkdir(exist_ok=True)
    try:
        data = json.loads(UNITS_PATH.read_text(encoding="utf-8")) if UNITS_PATH.exists() else {}
    except (OSError, json.JSONDecodeError):
        data = {}
    data[model_key] = units
    UNITS_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
