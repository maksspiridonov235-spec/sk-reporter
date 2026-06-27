"""Список файлов и метаданных для раздела «Планирование»."""

from __future__ import annotations

from typing import Any

from sk_reporter.personnel_store import is_engineer, load_people

_SECTIONS = frozenset({"personnel", "otkk", "contractors", "projects"})


def list_personnel() -> dict[str, Any]:
    from sk_reporter.personnel_db import db_status

    people = []
    engineers_count = 0
    for p in load_people():
        eng = is_engineer(p)
        if eng:
            engineers_count += 1
        people.append({**p, "is_engineer": eng})
    return {
        "storage": "postgresql",
        "people_count": len(people),
        "engineers_count": engineers_count,
        "people": people,
        "db": db_status(),
    }


def list_otkk() -> dict[str, Any]:
    from sk_reporter.otkk_db import db_status
    from sk_reporter.otkk_store import load_cards

    cards = []
    for card in load_cards():
        cards.append(
            {
                "id": card["id"],
                "code": card.get("code") or "",
                "title": card.get("title") or "",
                "source_file": card.get("file") or "",
            }
        )
    db = db_status()
    return {
        "storage": "postgresql",
        "count": len(cards),
        "cards": cards,
        "db": db,
    }


def list_contractors() -> dict[str, Any]:
    from sk_reporter.contractor_db import db_status, list_contractors as db_list

    contractors = db_list()
    return {
        "storage": "postgresql",
        "count": len(contractors),
        "contractors": contractors,
        "db": db_status(),
    }


def list_projects() -> dict[str, Any]:
    from sk_reporter.project_db import db_status, list_projects_catalog

    projects = list_projects_catalog()
    return {
        "storage": "postgresql",
        "count": len(projects),
        "projects": projects,
        "db": db_status(),
    }


def planning_section(section: str) -> dict[str, Any]:
    if section not in _SECTIONS:
        raise KeyError(section)
    if section == "personnel":
        return {"section": section, **list_personnel()}
    if section == "otkk":
        return {"section": section, **list_otkk()}
    if section == "contractors":
        return {"section": section, **list_contractors()}
    return {"section": section, **list_projects()}
