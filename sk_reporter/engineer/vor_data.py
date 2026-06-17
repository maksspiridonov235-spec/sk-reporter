"""Загрузка vor.json для UI."""

from __future__ import annotations

import json
from typing import Any

from sk_reporter.engineer.tk_catalog import resolve_tk_for_work
from sk_reporter.paths import project_dir


def load_project_meta(project_id: str) -> dict[str, Any]:
    return {"id": project_id, "title": project_id}


def load_vor_json(project_id: str) -> dict[str, Any]:
    path = project_dir(project_id) / "vor.json"
    if not path.is_file():
        raise FileNotFoundError(f"Нет vor.json для {project_id}. Запустите: python scripts/build_engineer_data.py --vor")
    return json.loads(path.read_text(encoding="utf-8"))


def flatten_vor_works(project_id: str) -> list[dict[str, Any]]:
    vor = load_vor_json(project_id)
    proj = project_dir(project_id)
    out: list[dict[str, Any]] = []
    idx = 0
    for stage in vor.get("stages") or []:
        stage_title = stage.get("title") or ""
        for obj in stage.get("objects") or []:
            obj_title = obj.get("title") or ""
            for work in obj.get("works") or []:
                name = work.get("name") or ""
                key = f"{stage_title}|{obj_title}|{name}"
                out.append(
                    {
                        "key": key,
                        "idx": idx,
                        "stage": stage_title,
                        "object": obj_title,
                        "name": name,
                        "unit": work.get("unit") or "",
                        "quantity": work.get("quantity") or "",
                        "tk_id": resolve_tk_for_work(name, proj) or "",
                    }
                )
                idx += 1
        for work in stage.get("works") or []:
            name = work.get("name") or ""
            key = f"{stage_title}||{name}"
            out.append(
                {
                    "key": key,
                    "idx": idx,
                    "stage": stage_title,
                    "object": "",
                    "name": name,
                    "unit": work.get("unit") or "",
                    "quantity": work.get("quantity") or "",
                    "tk_id": resolve_tk_for_work(name, proj) or "",
                }
            )
            idx += 1
    return out


def _project_ids_for_profile(profile: dict[str, Any]) -> list[str]:
    return []


def profile_project_ids(profile: dict[str, Any]) -> list[str]:
    return _project_ids_for_profile(profile)


def list_profile_projects(profile: dict[str, Any]) -> list[dict[str, Any]]:
    result = []
    for pid in _project_ids_for_profile(profile):
        meta = load_project_meta(pid)
        display = meta.get("title") or pid
        try:
            works = flatten_vor_works(pid)
        except FileNotFoundError:
            works = []
        result.append(
            {
                "id": pid,
                "title": display,
                "object_name": meta.get("object_name") or "",
                "works_count": len(works),
                "works": works,
            }
        )
    return result
