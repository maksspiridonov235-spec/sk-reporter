"""Хаб инженеров: карточки по назначениям на проекты, автопрофили."""

from __future__ import annotations

from typing import Any

import yaml

from sk_reporter.paths import engineer_profiles_dir, repo_root
from sk_reporter.personnel_store import get_person, list_engineers
from sk_reporter.project_store import engineer_project_map, get_project


def _scan_profiles_by_person() -> dict[str, str]:
    """person_id → profile id (имя yaml без расширения)."""
    out: dict[str, str] = {}
    root = engineer_profiles_dir()
    if not root.is_dir():
        return out
    for path in sorted(root.glob("*.yaml")):
        if path.stem == "example":
            continue
        data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        pid = str(data.get("person_id") or "").strip()
        if pid:
            out[pid] = data.get("id") or path.stem
    return out


def find_profile_id(person_id: str) -> str | None:
    return _scan_profiles_by_person().get(str(person_id).strip())


def ensure_engineer_profile(person_id: str) -> str:
    """Создать yaml при первом назначении; вернуть profile id."""
    person_id = str(person_id).strip()
    existing = find_profile_id(person_id)
    if existing:
        return existing

    person = get_person(person_id)
    if not person:
        raise KeyError(f"person_id «{person_id}» не найден в справочнике сотрудников")

    profile_id = person_id
    profile_path = engineer_profiles_dir() / f"{profile_id}.yaml"
    if not profile_path.is_file():
        profile_path.parent.mkdir(parents=True, exist_ok=True)
        profile_path.write_text(
            yaml.safe_dump(
                {
                    "id": profile_id,
                    "person_id": person_id,
                    "projects": [],
                    "report_template": "data/engineer/report_template.docx",
                },
                allow_unicode=True,
                sort_keys=False,
            ),
            encoding="utf-8",
        )

    return profile_id


def ensure_profiles_for_engineers(engineer_ids: list[str]) -> None:
    for eid in engineer_ids:
        try:
            ensure_engineer_profile(str(eid))
        except KeyError:
            continue


def list_hub_engineers() -> list[dict[str, Any]]:
    """Инженеры с хотя бы одним назначением на проект."""
    by_person = engineer_project_map()
    profile_by_person = _scan_profiles_by_person()
    items: list[dict[str, Any]] = []

    for person in list_engineers():
        pid = person["id"]
        projects_raw = by_person.get(pid) or []
        if not projects_raw:
            continue

        profile_id = profile_by_person.get(pid)
        if not profile_id:
            try:
                profile_id = ensure_engineer_profile(pid)
            except KeyError:
                profile_id = None
        projects = []
        for pr in projects_raw:
            rich = get_project(pr["id"]) or {}
            projects.append(
                {
                    "id": pr["id"],
                    "title": rich.get("object_name") or rich.get("title") or pr["title"],
                }
            )

        items.append(
            {
                "person_id": pid,
                "profile_id": profile_id,
                "fio": person["fio"],
                "position": person.get("position") or "",
                "projects": projects,
                "projects_count": len(projects),
                "profile_ok": bool(profile_id),
                "href": f"/engineer/{profile_id}" if profile_id else None,
            }
        )

    return sorted(items, key=lambda x: x["fio"].casefold())


def hub_payload() -> dict[str, Any]:
    engineers = list_hub_engineers()
    return {
        "engineers": engineers,
        "engineers_count": len(engineers),
        "profiles_dir": str(engineer_profiles_dir().relative_to(repo_root())),
    }
