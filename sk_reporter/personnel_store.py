"""Справочник персонала в PostgreSQL (RelaxDev, DATABASE_URL)."""

from __future__ import annotations

import re
from typing import Any

from sk_reporter.db.config import database_enabled

_ENGINEER_MARKERS = ("инженер ск", "инженер строительного")


def person_id_from_fio(fio: str) -> str:
    parts = fio.strip().split()
    if not parts:
        return "unknown"
    surname = parts[0].lower()
    suffix = parts[1][0].lower() if len(parts) > 1 else ""
    raw = f"{surname}-{suffix}" if suffix else surname
    return re.sub(r"[^\w\-а-яё]", "", raw, flags=re.I)


def _normalize_fio(fio: str) -> str:
    return " ".join(fio.split())


def _require_database() -> None:
    if not database_enabled():
        raise RuntimeError(
            "DATABASE_URL не задан — справочник сотрудников хранится только в PostgreSQL"
        )


def load_people() -> list[dict[str, Any]]:
    _require_database()
    from sk_reporter.personnel_db import load_people_from_db

    return load_people_from_db()


def list_engineers() -> list[dict[str, Any]]:
    result = []
    for p in load_people():
        if is_engineer(p):
            result.append(p)
    return sorted(result, key=lambda x: x["fio"])


def is_engineer(person: dict[str, Any]) -> bool:
    pos = (person.get("position") or "").lower()
    return any(m in pos for m in _ENGINEER_MARKERS)


def get_person(person_id: str) -> dict[str, Any] | None:
    for p in load_people():
        if p["id"] == person_id:
            return p
    return None
