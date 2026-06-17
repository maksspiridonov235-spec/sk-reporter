"""Проекты в PostgreSQL: карточка с ВОР и ТЛ (как ОТКК)."""

from __future__ import annotations

from typing import Any

from sqlalchemy import text

from sk_reporter.db.config import database_enabled, database_url
from sk_reporter.db.models import Contractor, Project, ProjectEngineer
from sk_reporter.db.session import get_session, init_db
from sk_reporter.paths import projects_dir
from sk_reporter.personnel_store import get_person, list_engineers
from sk_reporter.project_import import import_project_folder


def _require_database() -> None:
    if not database_enabled():
        raise RuntimeError("DATABASE_URL не задан — проекты хранятся только в PostgreSQL")


def _ensure_project_schema() -> None:
    from sqlalchemy import create_engine

    url = database_url()
    if not url:
        return
    engine = create_engine(url, pool_pre_ping=True)
    stmts = [
        "ALTER TABLE projects ADD COLUMN IF NOT EXISTS vor_file VARCHAR(512) NOT NULL DEFAULT ''",
        "ALTER TABLE projects ADD COLUMN IF NOT EXISTS tl_file VARCHAR(512) NOT NULL DEFAULT ''",
        "ALTER TABLE projects ADD COLUMN IF NOT EXISTS content JSONB",
        "ALTER TABLE projects ALTER COLUMN contractor_id DROP NOT NULL",
    ]
    with engine.begin() as conn:
        for stmt in stmts:
            conn.execute(text(stmt))


def _count_vor_works(vor: dict[str, Any] | None) -> int:
    if not vor:
        return 0
    total = 0
    for stage in vor.get("stages") or []:
        total += len(stage.get("works") or [])
        for obj in stage.get("objects") or []:
            total += len(obj.get("works") or [])
    return total


def _catalog_row(project: Project, *, include_content: bool = False) -> dict[str, Any]:
    content = project.content or {}
    vor = content.get("vor") or {}
    tl = content.get("tl") or {}
    out: dict[str, Any] = {
        "id": project.id,
        "title": project.title or project.id,
        "object_name": project.object_name or "",
        "vor_file": project.vor_file or "",
        "tl_file": project.tl_file or "",
        "has_content": project.content is not None,
        "vor_works_count": _count_vor_works(vor),
        "tl_tables_count": len(tl.get("tables") or []),
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


def upsert_project_imported(payload: dict[str, Any]) -> dict[str, Any]:
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


def seed_project_from_disk(project_id: str, *, overwrite: bool = False) -> dict[str, Any]:
    pid = str(project_id).strip()
    existing = get_project_catalog(pid, include_content=True)
    if existing and existing.get("has_content") and not overwrite and not _project_needs_reseed(existing):
        return {"id": pid, "skipped": True, "reason": "already in db"}

    payload = import_project_folder(pid)
    result = upsert_project_imported(payload)
    vor = (payload["content"].get("vor") or {}) if payload.get("content") else {}
    result["seeded"] = True
    result["vor_works_count"] = _count_vor_works(vor)
    return result


def seed_projects_from_disk(*, overwrite: bool = False) -> dict[str, Any]:
    root = projects_dir()
    seeded: list[str] = []
    skipped: list[dict[str, str]] = []
    if not root.is_dir():
        return {"seeded": seeded, "skipped": [{"id": "", "reason": "no projects dir"}]}
    for entry in sorted(root.iterdir(), key=lambda p: p.name.casefold()):
        if not entry.is_dir() or entry.name.startswith("."):
            continue
        try:
            result = seed_project_from_disk(entry.name, overwrite=overwrite)
            if result.get("seeded"):
                seeded.append(entry.name)
            elif result.get("skipped"):
                skipped.append({"id": entry.name, "reason": result.get("reason", "skip")})
        except Exception as exc:
            skipped.append({"id": entry.name, "reason": str(exc)})
    return {"seeded": seeded, "seeded_count": len(seeded), "skipped": skipped}


def seed_sup_pdr_pilot(*, overwrite: bool = True) -> dict[str, Any]:
    """Пилот SUP-PDR: ВОР из эталона в репо (как ОТКК), не зависит от .doc на сервере."""
    from sk_reporter.project_etalon import sup_pdr_enc_00_1_payload

    pid = "SUP-PDR-ENC-001-DD-ST01-EV_00.1"
    existing = get_project_catalog(pid, include_content=True)
    if existing and existing.get("has_content") and not overwrite and not _project_needs_reseed(existing):
        return {"id": pid, "skipped": True, "reason": "already in db"}

    payload = sup_pdr_enc_00_1_payload()
    result = upsert_project_imported(payload)
    vor = (payload["content"].get("vor") or {}) if payload.get("content") else {}
    result["seeded"] = True
    result["vor_works_count"] = _count_vor_works(vor)
    result["source"] = "etalon"
    return result


# --- назначения инженеров (позже, без изменений API) ---

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


def list_projects_rich(*, contractor_id: str | None = None) -> list[dict[str, Any]]:
    _require_database()
    init_db()
    _ensure_project_schema()
    with get_session() as session:
        q = session.query(Project).filter(Project.is_active.is_(True))
        if contractor_id:
            q = q.filter(Project.contractor_id == contractor_id)
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


def set_project_engineers(project_id: str, engineer_ids: list[str]) -> dict[str, Any]:
    _require_database()
    init_db()
    valid_ids = {e["id"] for e in list_engineers()}
    cleaned = [str(eid).strip() for eid in engineer_ids if str(eid).strip() in valid_ids]
    with get_session() as session:
        project = session.get(Project, project_id)
        if not project or not project.is_active:
            raise FileNotFoundError(f"Проект не найден: {project_id}")
        session.query(ProjectEngineer).filter(ProjectEngineer.project_id == project_id).delete()
        for eid in cleaned:
            session.add(ProjectEngineer(project_id=project_id, person_id=eid))
        session.flush()
        contractor = session.get(Contractor, project.contractor_id) if project.contractor_id else None
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
