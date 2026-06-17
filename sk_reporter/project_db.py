"""Проекты и назначения инженеров в PostgreSQL."""

from __future__ import annotations

from typing import Any

from sk_reporter.db.config import database_enabled
from sk_reporter.db.models import Contractor, Project, ProjectEngineer
from sk_reporter.db.session import get_session, init_db
from sk_reporter.personnel_store import get_person, list_engineers


def _require_database() -> None:
    if not database_enabled():
        raise RuntimeError("DATABASE_URL не задан — проекты хранятся только в PostgreSQL")


def db_status() -> dict[str, Any]:
    if not database_enabled():
        return {
            "enabled": False,
            "configured": False,
            "count": 0,
            "ok": False,
            "error": "DATABASE_URL не задан",
        }
    try:
        init_db()
        with get_session() as session:
            count = session.query(Project).filter(Project.is_active.is_(True)).count()
        return {"enabled": True, "configured": True, "count": count, "ok": True}
    except Exception as exc:
        return {"enabled": True, "configured": True, "count": 0, "ok": False, "error": str(exc)}


def _resolve_engineers(ids: list[str]) -> list[dict[str, Any]]:
    resolved = []
    for eid in ids:
        eid = str(eid).strip()
        if not eid:
            continue
        person = get_person(eid)
        if person:
            resolved.append(person)
        else:
            resolved.append({"id": eid, "fio": eid, "phone": "", "position": ""})
    return resolved


def _project_row_to_dict(
    project: Project,
    contractor: Contractor | None,
    engineer_ids: list[str],
) -> dict[str, Any]:
    engineers = _resolve_engineers(engineer_ids)
    title = project.title or project.id
    object_name = project.object_name or title
    return {
        "id": project.id,
        "contractor_id": project.contractor_id,
        "contractor_name": contractor.name if contractor else "",
        "title": title,
        "object_name": object_name,
        "is_active": bool(project.is_active),
        "engineer_ids": engineer_ids,
        "engineers": engineers,
        "vor": {
            "ready": False,
            "message": "ВОР в БД — позже; файлы vor.json пока на диске",
        },
        "tk_mappings": 0,
    }


def list_projects_rich(*, contractor_id: str | None = None) -> list[dict[str, Any]]:
    _require_database()
    init_db()
    with get_session() as session:
        q = session.query(Project).filter(Project.is_active.is_(True))
        if contractor_id:
            q = q.filter(Project.contractor_id == contractor_id)
        projects = q.order_by(Project.title, Project.id).all()
        if not projects:
            return []
        pids = [p.id for p in projects]
        contractor_ids = {p.contractor_id for p in projects}
        contractors = {
            c.id: c
            for c in session.query(Contractor).filter(Contractor.id.in_(contractor_ids)).all()
        }
        links = (
            session.query(ProjectEngineer)
            .filter(ProjectEngineer.project_id.in_(pids))
            .all()
        )
        by_project: dict[str, list[str]] = {pid: [] for pid in pids}
        for link in links:
            by_project.setdefault(link.project_id, []).append(link.person_id)
        return [
            _project_row_to_dict(p, contractors.get(p.contractor_id), by_project.get(p.id, []))
            for p in projects
        ]


def get_project(project_id: str) -> dict[str, Any] | None:
    _require_database()
    init_db()
    with get_session() as session:
        project = session.get(Project, project_id)
        if not project or not project.is_active:
            return None
        contractor = session.get(Contractor, project.contractor_id)
        engineer_ids = [
            r.person_id
            for r in session.query(ProjectEngineer)
            .filter(ProjectEngineer.project_id == project_id)
            .all()
        ]
        return _project_row_to_dict(project, contractor, engineer_ids)


def create_project(
    project_id: str,
    *,
    contractor_id: str,
    title: str = "",
    object_name: str = "",
) -> dict[str, Any]:
    _require_database()
    pid = str(project_id).strip()
    cid = str(contractor_id).strip()
    if not pid:
        raise ValueError("Код проекта обязателен")
    if not cid:
        raise ValueError("Подрядчик обязателен")
    init_db()
    with get_session() as session:
        if not session.get(Contractor, cid):
            raise KeyError(f"Подрядчик не найден: {cid}")
        if session.get(Project, pid):
            raise ValueError(f"Проект уже существует: {pid}")
        title = (title or pid).strip()
        object_name = (object_name or title).strip()
        row = Project(
            id=pid,
            contractor_id=cid,
            title=title,
            object_name=object_name,
            is_active=True,
        )
        session.add(row)
        session.flush()
        contractor = session.get(Contractor, cid)
        return _project_row_to_dict(row, contractor, [])


def set_project_engineers(project_id: str, engineer_ids: list[str]) -> dict[str, Any]:
    _require_database()
    init_db()
    valid_ids = {e["id"] for e in list_engineers()}
    cleaned = []
    for eid in engineer_ids:
        eid = str(eid).strip()
        if eid and eid in valid_ids:
            cleaned.append(eid)
    with get_session() as session:
        project = session.get(Project, project_id)
        if not project or not project.is_active:
            raise FileNotFoundError(f"Проект не найден: {project_id}")
        session.query(ProjectEngineer).filter(ProjectEngineer.project_id == project_id).delete()
        for eid in cleaned:
            session.add(ProjectEngineer(project_id=project_id, person_id=eid))
        session.flush()
        contractor = session.get(Contractor, project.contractor_id)
        return _project_row_to_dict(project, contractor, cleaned)


def engineer_project_map() -> dict[str, list[dict[str, str]]]:
    _require_database()
    init_db()
    out: dict[str, list[dict[str, str]]] = {}
    for proj in list_projects_rich():
        label = proj.get("object_name") or proj.get("title") or proj["id"]
        for eid in proj.get("engineer_ids") or []:
            out.setdefault(str(eid), []).append({"id": proj["id"], "title": label})
    return out
