"""Профиль инженера из справочника PostgreSQL."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

from sk_reporter.paths import repo_root
from sk_reporter.personnel_store import get_person, is_engineer

DEFAULT_REPORT_TEMPLATE = "data/engineer/report_template.docx"


def load_profile(person_id: str | None = None) -> dict[str, Any]:
    pid = (person_id or os.environ.get("SK_ENGINEER_PROFILE", "")).strip()
    if not pid:
        raise ValueError("Не задан инженер: параметр person_id или SK_ENGINEER_PROFILE")

    person = get_person(pid)
    if not person:
        raise ValueError(f"Сотрудник не найден в справочнике: {pid}")
    if not is_engineer(person):
        raise ValueError(f"«{person['fio']}» не является инженером СК")

    return {
        "id": pid,
        "person_id": pid,
        "name": person["fio"],
        "position": person.get("position") or "",
        "phone": person.get("phone") or "",
        "report_template": DEFAULT_REPORT_TEMPLATE,
    }


def resolve_report_template(profile: dict[str, Any]) -> Path | None:
    raw = profile.get("report_template") or DEFAULT_REPORT_TEMPLATE
    p = Path(raw)
    if not p.is_absolute():
        p = repo_root() / p
    return p if p.is_file() else None
