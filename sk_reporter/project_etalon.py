"""Эталоны карточек проектов в репозитории → seed в PostgreSQL (как ОТКК)."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

_DATA_DIR = Path(__file__).resolve().parent / "project_data"


def _load_json(name: str) -> dict[str, Any]:
    path = _DATA_DIR / name
    if not path.is_file():
        raise FileNotFoundError(f"Нет эталона проекта: {path}")
    return json.loads(path.read_text(encoding="utf-8"))


def sup_pdr_enc_00_1_payload() -> dict[str, Any]:
    return _load_json("sup_pdr_enc_00_1.json")
