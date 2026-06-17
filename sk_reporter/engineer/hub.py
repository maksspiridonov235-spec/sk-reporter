"""Хаб инженеров: карточки по назначениям на проекты (данные из PostgreSQL)."""

from __future__ import annotations

from typing import Any

from sk_reporter.personnel_store import list_engineers
from sk_reporter.project_store import engineer_project_map, get_project


def list_hub_engineers() -> list[dict[str, Any]]:
    """Инженеры с хотя бы одним назначением на проект."""
    by_person = engineer_project_map()
    items: list[dict[str, Any]] = []

    for person in list_engineers():
        pid = person["id"]
        projects_raw = by_person.get(pid) or []
        if not projects_raw:
            continue

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
                "profile_id": pid,
                "fio": person["fio"],
                "position": person.get("position") or "",
                "projects": projects,
                "projects_count": len(projects),
                "profile_ok": True,
                "href": f"/engineer/{pid}",
            }
        )

    return sorted(items, key=lambda x: x["fio"].casefold())


def hub_payload() -> dict[str, Any]:
    engineers = list_hub_engineers()
    return {
        "engineers": engineers,
        "engineers_count": len(engineers),
    }
