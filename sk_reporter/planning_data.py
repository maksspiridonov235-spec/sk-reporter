"""Список файлов и метаданных для раздела «Планирование»."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml

from sk_reporter.luvr_store import luvr_planning_payload
from sk_reporter.personnel_store import is_engineer, list_engineers, load_people
from sk_reporter.project_store import engineer_project_map, get_project, list_projects_rich, set_project_engineers
from sk_reporter.paths import personnel_dir, repo_root, tk_dir

_SECTIONS = frozenset({"projects", "luvr", "personnel", "otkk"})


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


def list_luvr() -> dict[str, Any]:
    return luvr_planning_payload()


def list_personnel() -> dict[str, Any]:
    folder = personnel_dir()
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
        "folder": str(folder.relative_to(repo_root())),
        "people_count": len(people),
        "engineers_count": engineers_count,
        "people": people,
    }


def list_otkk() -> dict[str, Any]:
    folder = tk_dir()
    cards = []
    manifest = folder / "manifest.yaml"
    if manifest.is_file():
        data = yaml.safe_load(manifest.read_text(encoding="utf-8")) or {}
        for card in data.get("cards") or []:
            fname = card.get("file") or ""
            fp = folder / fname
            cards.append(
                {
                    "id": card.get("id"),
                    "file": fname,
                    "present": fp.is_file(),
                    "size_kb": round(fp.stat().st_size / 1024, 1) if fp.is_file() else None,
                }
            )
    else:
        for p in sorted(folder.iterdir()):
            if p.suffix.lower() in {".doc", ".docx"} and not p.name.startswith("."):
                cards.append(
                    {
                        "id": p.stem[:20],
                        "file": p.name,
                        "present": True,
                        "size_kb": round(p.stat().st_size / 1024, 1),
                    }
                )
    return {
        "folder": str(folder.relative_to(repo_root())),
        "count": len(cards),
        "cards": cards,
    }


def planning_section(section: str) -> dict[str, Any]:
    if section not in _SECTIONS:
        raise KeyError(section)
    if section == "projects":
        return projects_planning_payload()
    if section == "luvr":
        return {"section": section, **list_luvr()}
    if section == "personnel":
        return {"section": section, **list_personnel()}
    return {"section": section, **list_otkk()}
