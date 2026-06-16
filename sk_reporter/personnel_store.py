"""Справочник персонала: PostgreSQL (если DATABASE_URL) или personnel.yaml."""

from __future__ import annotations

import re
from functools import lru_cache
from typing import Any

import yaml

from sk_reporter.db.config import database_enabled
from sk_reporter.paths import personnel_dir

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


def storage_backend() -> str:
    return "postgresql" if database_enabled() else "yaml"


@lru_cache(maxsize=1)
def _load_people_from_yaml() -> list[dict[str, Any]]:
    path = personnel_dir() / "personnel.yaml"
    if not path.is_file():
        return []
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    out: list[dict[str, Any]] = []
    seen: set[str] = set()
    for row in data.get("people") or []:
        fio = _normalize_fio(str(row.get("ФИО") or ""))
        if not fio:
            continue
        pid = str(row.get("id") or person_id_from_fio(fio))
        if pid in seen:
            pid = f"{pid}-{len(seen)}"
        seen.add(pid)
        out.append(
            {
                "id": pid,
                "fio": fio,
                "phone": str(row.get("Телефон") or "").strip(),
                "position": str(row.get("Должность") or "").strip(),
                "control_mode": str(row.get("Режим контроля") or "").strip(),
            }
        )
    return out


def clear_personnel_cache() -> None:
    _load_people_from_yaml.cache_clear()


def load_people() -> list[dict[str, Any]]:
    if database_enabled():
        from sk_reporter.personnel_db import load_people_from_db

        return load_people_from_db()
    return _load_people_from_yaml()


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
