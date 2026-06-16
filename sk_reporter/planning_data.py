"""Список файлов и метаданных для раздела «Планирование»."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from sk_reporter.personnel_store import is_engineer, list_engineers, load_people
from sk_reporter.project_store import engineer_project_map, get_project, list_projects_rich
from sk_reporter.paths import repo_root, tk_dir

_SECTIONS = frozenset({"projects", "personnel", "otkk"})


def _file_row(path: Path) -> dict[str, Any]:
    st = path.stat()
    root = repo_root()
    try:
        rel = str(path.relative_to(root))
    except ValueError:
        rel = str(path)
    return {
        "name": path.name,
        "rel": rel,
        "size_kb": round(st.st_size / 1024, 1),
        "suffix": path.suffix.lower(),
    }


def _list_files(folder: Path, pattern: str = "*") -> list[dict[str, Any]]:
    if not folder.is_dir():
        return []
    rows = []
    for p in sorted(folder.glob(pattern)):
        if p.name.startswith("."):
            continue
        if p.is_file():
            rows.append(_file_row(p))
    return rows


def list_projects() -> list[dict[str, Any]]:
    return list_projects_rich()


def projects_planning_payload() -> dict[str, Any]:
    from sk_reporter.deployment_store import deployment_status

    dep = deployment_status()
    return {
        "section": "projects",
        "engineers_available": list_engineers(),
        "items": list_projects(),
        "deployment": dep,
        "assignment_stats": dep.get("assignments") or {},
    }


def list_personnel() -> dict[str, Any]:
    from sk_reporter.personnel_db import db_status

    assignments = engineer_project_map()
    people = []
    engineers_count = 0
    for p in load_people():
        eng = is_engineer(p)
        if eng:
            engineers_count += 1
        people.append(
            {
                **p,
                "is_engineer": eng,
                "projects": assignments.get(p["id"], []),
            }
        )
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

    folder = tk_dir()
    cards = []
    for card in load_cards():
        fp = folder / card["file"]
        cards.append(
            {
                "id": card["id"],
                "file": card["file"],
                "code": card.get("code") or "",
                "title": card.get("title") or "",
                "has_content": bool(card.get("has_content")),
                "present": fp.is_file(),
                "size_kb": round(fp.stat().st_size / 1024, 1) if fp.is_file() else None,
            }
        )
    present_count = sum(1 for c in cards if c["present"])
    content_count = sum(1 for c in cards if c["has_content"])
    return {
        "storage": "postgresql",
        "folder": str(folder.relative_to(repo_root())),
        "count": len(cards),
        "present_count": present_count,
        "content_count": content_count,
        "cards": cards,
        "db": db_status(),
    }


def planning_section(section: str) -> dict[str, Any]:
    if section not in _SECTIONS:
        raise KeyError(section)
    if section == "projects":
        return projects_planning_payload()
    if section == "personnel":
        return {"section": section, **list_personnel()}
    return {"section": section, **list_otkk()}
