"""Эталоны карточек проектов в репозитории → seed в PostgreSQL (как ОТКК)."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from sk_reporter.project_sup_pdr_tl_data import sup_pdr_tl_card_fields, sup_pdr_tl_content

_DATA_DIR = Path(__file__).resolve().parent / "project_data"


def _load_json(name: str) -> dict[str, Any]:
    path = _DATA_DIR / name
    if not path.is_file():
        raise FileNotFoundError(f"Нет эталона проекта: {path}")
    return json.loads(path.read_text(encoding="utf-8"))


def sup_pdr_enc_00_1_payload() -> dict[str, Any]:
    payload = _load_json("sup_pdr_enc_00_1.json")
    tl_meta = sup_pdr_tl_card_fields()
    payload["title"] = tl_meta["title"]
    payload["object_name"] = tl_meta["object_name"]
    payload["tl_file"] = tl_meta["tl_file"]
    content = payload.setdefault("content", {})
    content["tl"] = sup_pdr_tl_content()
    return payload
