"""Проекты в PostgreSQL: карточка с ВОР и ТЛ (как ОТКК)."""

from __future__ import annotations

from typing import Any

from sk_reporter.db.config import database_enabled
from sk_reporter.db.models import Contractor, Project, ProjectEngineer
from sk_reporter.db.schema_ensure import ensure_table_columns
from sk_reporter.db.session import get_session, init_db
from sk_reporter.personnel_store import get_person


def _require_database() -> None:
    if not database_enabled():
        raise RuntimeError("DATABASE_URL не задан — проекты хранятся только в PostgreSQL")


def _ensure_project_schema() -> None:
    ensure_table_columns(
        "projects",
        add_columns={
            "vor_file": "VARCHAR(512) NOT NULL DEFAULT ''",
            "tl_file": "VARCHAR(512) NOT NULL DEFAULT ''",
            "content": "JSONB",
        },
        drop_not_null=["contractor_id"],
    )


def _count_vor_works(vor: dict[str, Any] | None) -> int:
    if not vor:
        return 0
    rows = vor.get("rows")
    if rows:
        return sum(1 for row in rows if str(row.get("kind") or "") == "work")
    total = 0
    for stage in vor.get("stages") or []:
        total += len(stage.get("works") or [])
        for obj in stage.get("objects") or []:
            total += len(obj.get("works") or [])
    return total


def _tl_cipher(content: dict[str, Any]) -> str:
    tl = content.get("tl") or {}
    for row in tl.get("rows") or []:
        if str(row.get("label") or "").strip() == "Шифр проекта":
            return str(row.get("value") or "").strip()
    return ""


def _catalog_row(project: Project, *, include_content: bool = False) -> dict[str, Any]:
    content = project.content or {}
    vor = content.get("vor") or {}
    tl = content.get("tl") or {}
    cipher = _tl_cipher(content) or project.id
    out: dict[str, Any] = {
        "id": project.id,
        "cipher": cipher,
        "title": project.title or project.id,
        "object_name": project.object_name or "",
        "vor_file": project.vor_file or "",
        "tl_file": project.tl_file or "",
        "has_content": project.content is not None,
        "vor_works_count": _count_vor_works(vor),
        "tl_tables_count": len(tl.get("tables") or []) or len(tl.get("rows") or []),
        "contractor_id": project.contractor_id or "",
    }
    if include_content and project.content is not None:
        out["content"] = project.content
    return out


def db_status() -> dict[str, Any]:
    if not database_enabled():
        return {
            "enabled": False,
            "configured": False,
            "count": 0,
            "with_content": 0,
            "ok": False,
            "error": "DATABASE_URL не задан",
        }
    try:
        init_db()
        _ensure_project_schema()
        with get_session() as session:
            active = session.query(Project).filter(Project.is_active.is_(True)).count()
            with_content = (
                session.query(Project)
                .filter(Project.is_active.is_(True), Project.content.isnot(None))
                .count()
            )
        return {
            "enabled": True,
            "configured": True,
            "count": active,
            "with_content": with_content,
            "ok": True,
        }
    except Exception as exc:
        return {
            "enabled": True,
            "configured": True,
            "count": 0,
            "with_content": 0,
            "ok": False,
            "error": str(exc),
        }


def list_projects_catalog() -> list[dict[str, Any]]:
    _require_database()
    init_db()
    _ensure_project_schema()
    with get_session() as session:
        rows = (
            session.query(Project)
            .filter(Project.is_active.is_(True))
            .order_by(Project.title, Project.id)
            .all()
        )
        return [_catalog_row(r) for r in rows]


def get_project_catalog(project_id: str, *, include_content: bool = False) -> dict[str, Any] | None:
    pid = str(project_id).strip()
    if not pid:
        return None
    _require_database()
    init_db()
    _ensure_project_schema()
    with get_session() as session:
        row = session.get(Project, pid)
        if not row or not row.is_active:
            return None
        return _catalog_row(row, include_content=include_content)


def get_project_vor_content(project_id: str) -> dict[str, Any] | None:
    card = get_project_catalog(project_id, include_content=True)
    if not card or not card.get("content"):
        return None
    vor = (card["content"] or {}).get("vor")
    return vor if isinstance(vor, dict) else None


def upsert_project_card(payload: dict[str, Any]) -> dict[str, Any]:
    pid = str(payload.get("id") or "").strip()
    if not pid:
        raise ValueError("Нет id проекта")
    content = payload.get("content")
    if not isinstance(content, dict):
        raise ValueError("Нет content проекта")

    _require_database()
    init_db()
    _ensure_project_schema()
    with get_session() as session:
        row = session.get(Project, pid)
        fields = {
            "title": str(payload.get("title") or pid).strip(),
            "object_name": str(payload.get("object_name") or pid).strip(),
            "vor_file": str(payload.get("vor_file") or "").strip(),
            "tl_file": str(payload.get("tl_file") or "").strip(),
            "content": content,
            "is_active": True,
        }
        if row:
            for key, val in fields.items():
                setattr(row, key, val)
        else:
            row = Project(id=pid, contractor_id=None, **fields)
            session.add(row)
        session.flush()
        return _catalog_row(row)


def _project_needs_reseed(existing: dict[str, Any] | None) -> bool:
    if not existing or not existing.get("has_content"):
        return True
    return (existing.get("vor_works_count") or 0) == 0


def seed_project_from_etalon(payload: dict[str, Any], *, overwrite: bool = True) -> dict[str, Any]:
    pid = str(payload.get("id") or "").strip()
    if not pid:
        raise ValueError("Нет id проекта")
    existing = get_project_catalog(pid, include_content=True)
    if existing and existing.get("has_content") and not overwrite and not _project_needs_reseed(existing):
        return {"id": pid, "skipped": True, "reason": "already in db"}

    result = upsert_project_card(payload)
    vor = (payload.get("content") or {}).get("vor") or {}
    result["seeded"] = True
    result["vor_works_count"] = _count_vor_works(vor)
    result["source"] = "etalon"
    return result


def seed_projects_pilots(*, overwrite: bool = True) -> dict[str, Any]:
    from sk_reporter.project_etalon import all_project_etalon_payloads

    seeded: list[str] = []
    skipped: list[dict[str, str]] = []
    for payload in all_project_etalon_payloads():
        pid = str(payload.get("id") or "")
        try:
            result = seed_project_from_etalon(payload, overwrite=overwrite)
            if result.get("seeded"):
                seeded.append(pid)
            elif result.get("skipped"):
                skipped.append({"id": pid, "reason": str(result.get("reason", "skip"))})
        except Exception as exc:
            skipped.append({"id": pid, "reason": str(exc)})
    return {"seeded": seeded, "seeded_count": len(seeded), "skipped": skipped}


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
    base = _catalog_row(project)
    base.update(
        {
            "contractor_name": contractor.name if contractor else "",
            "is_active": bool(project.is_active),
            "engineer_ids": engineer_ids,
            "engineers": _resolve_engineers(engineer_ids),
        }
    )
    return base


def list_projects_rich() -> list[dict[str, Any]]:
    _require_database()
    init_db()
    _ensure_project_schema()
    with get_session() as session:
        q = session.query(Project).filter(Project.is_active.is_(True))
        projects = q.order_by(Project.title, Project.id).all()
        if not projects:
            return []
        pids = [p.id for p in projects]
        contractor_ids = {p.contractor_id for p in projects if p.contractor_id}
        contractors = (
            {
                c.id: c
                for c in session.query(Contractor).filter(Contractor.id.in_(contractor_ids)).all()
            }
            if contractor_ids
            else {}
        )
        links = session.query(ProjectEngineer).filter(ProjectEngineer.project_id.in_(pids)).all()
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
    _ensure_project_schema()
    with get_session() as session:
        project = session.get(Project, project_id)
        if not project or not project.is_active:
            return None
        contractor = session.get(Contractor, project.contractor_id) if project.contractor_id else None
        engineer_ids = [
            r.person_id
            for r in session.query(ProjectEngineer)
            .filter(ProjectEngineer.project_id == project_id)
            .all()
        ]
        return _project_row_to_dict(project, contractor, engineer_ids)


def engineer_project_map() -> dict[str, list[dict[str, str]]]:
    _require_database()
    init_db()
    out: dict[str, list[dict[str, str]]] = {}
    for proj in list_projects_rich():
        label = proj.get("object_name") or proj.get("title") or proj["id"]
        for eid in proj.get("engineer_ids") or []:
            out.setdefault(str(eid), []).append({"id": proj["id"], "title": label})
    return out
