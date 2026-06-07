"""Проекты: метаданные, статистика ВОР, привязка инженеров."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import yaml

from sk_reporter.personnel_store import get_person, list_engineers
from sk_reporter.paths import project_dir, projects_dir, repo_root
from sk_reporter.project_title import resolve_object_name


def _read_project_yaml(proj: Path) -> dict[str, Any]:
    meta_path = proj / "project.yaml"
    meta: dict[str, Any] = {"id": proj.name, "title": proj.name}
    if meta_path.is_file():
        meta.update(yaml.safe_load(meta_path.read_text(encoding="utf-8")) or {})
    meta.setdefault("id", proj.name)
    meta.setdefault("title", proj.name)
    meta.setdefault("engineers", [])
    return meta


def _vor_stats(proj: Path) -> dict[str, Any]:
    cache = proj / "vor.json"
    if not cache.is_file():
        return {
            "ready": False,
            "stages": 0,
            "objects": 0,
            "works": 0,
            "message": "Нет vor.json — python scripts/build_engineer_data.py --vor",
        }
    data = json.loads(cache.read_text(encoding="utf-8"))
    stages = data.get("stages") or []
    objects = sum(len(s.get("objects") or []) for s in stages)
    works = sum(len(o.get("works") or []) for s in stages for o in (s.get("objects") or []))
    works += sum(len(s.get("works") or []) for s in stages)
    return {
        "ready": True,
        "source": data.get("source"),
        "stages": len(stages),
        "objects": objects,
        "works": works,
        "message": None,
    }


def _tk_map_count(proj: Path) -> int:
    path = proj / "work_tk_map.yaml"
    if not path.is_file():
        return 0
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    return len(data.get("mappings") or {})


def _resolve_engineers(ids: list[str]) -> list[dict[str, Any]]:
    resolved = []
    for eid in ids:
        if isinstance(eid, dict):
            eid = eid.get("id") or ""
        eid = str(eid).strip()
        if not eid:
            continue
        person = get_person(eid)
        if person:
            resolved.append(person)
        else:
            resolved.append({"id": eid, "fio": eid, "phone": "", "position": ""})
    return resolved


def engineer_project_map() -> dict[str, list[dict[str, str]]]:
    """person_id → список проектов, где назначен."""
    out: dict[str, list[dict[str, str]]] = {}
    for proj in list_projects_rich():
        for eid in proj.get("engineer_ids") or []:
            out.setdefault(str(eid), []).append(
                {
                    "id": proj["id"],
                    "title": proj.get("object_name") or proj["title"],
                }
            )
    return out


def get_project(project_id: str) -> dict[str, Any] | None:
    proj = project_dir(project_id)
    if not proj.is_dir():
        return None
    meta = _read_project_yaml(proj)
    engineer_ids = meta.get("engineers") or []
    if engineer_ids and isinstance(engineer_ids[0], dict):
        engineer_ids = [e.get("id") for e in engineer_ids if e.get("id")]
    parsed_name, title_page = resolve_object_name(proj, meta)
    return {
        "id": meta["id"],
        "title": meta.get("title") or meta["id"],
        "object_name": parsed_name,
        "title_page": title_page,
        "path": str(proj.relative_to(repo_root())),
        "vor_docx": meta.get("vor_docx"),
        "vor_doc": meta.get("vor_doc") or [],
        "vor": _vor_stats(proj),
        "tk_mappings": _tk_map_count(proj),
        "engineer_ids": engineer_ids,
        "engineers": _resolve_engineers(engineer_ids),
    }


def list_projects_rich() -> list[dict[str, Any]]:
    root = projects_dir()
    if not root.is_dir():
        return []
    items = []
    for proj in sorted(root.iterdir()):
        if not proj.is_dir() or proj.name.startswith("."):
            continue
        item = get_project(proj.name)
        if item:
            items.append(item)
    return items


def set_project_engineers(project_id: str, engineer_ids: list[str]) -> dict[str, Any]:
    proj = project_dir(project_id)
    if not proj.is_dir():
        raise FileNotFoundError(f"Проект не найден: {project_id}")

    meta_path = proj / "project.yaml"
    meta = _read_project_yaml(proj)
    valid_ids = {e["id"] for e in list_engineers()}
    cleaned = []
    for eid in engineer_ids:
        eid = str(eid).strip()
        if eid and eid in valid_ids:
            cleaned.append(eid)

    meta["engineers"] = cleaned
    meta_path.write_text(
        yaml.safe_dump(meta, allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )
    from sk_reporter.engineer.hub import ensure_profiles_for_engineers

    ensure_profiles_for_engineers(cleaned)
    result = get_project(project_id)
    if not result:
        raise RuntimeError("Не удалось прочитать проект после сохранения")
    return result
