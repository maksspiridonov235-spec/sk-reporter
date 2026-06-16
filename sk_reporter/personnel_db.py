"""Справочник сотрудников в PostgreSQL."""

from __future__ import annotations

from typing import Any

from sk_reporter.db.config import database_enabled
from sk_reporter.db.models import Personnel
from sk_reporter.db.session import get_session, init_db
from sk_reporter.personnel_store import _normalize_fio


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
            count = session.query(Personnel).count()
        return {"enabled": True, "configured": True, "count": count, "ok": True}
    except Exception as exc:
        return {"enabled": True, "configured": True, "count": 0, "ok": False, "error": str(exc)}


def _row_to_dict(row: Personnel) -> dict[str, Any]:
    return {
        "id": row.id,
        "fio": row.fio,
        "phone": row.phone or "",
        "position": row.position or "",
        "control_mode": row.control_mode or "",
    }


def load_people_from_db() -> list[dict[str, Any]]:
    init_db()
    with get_session() as session:
        rows = session.query(Personnel).order_by(Personnel.fio).all()
        return [_row_to_dict(r) for r in rows]


def upsert_people(people: list[dict[str, Any]]) -> dict[str, Any]:
    if not people:
        return {"upserted": 0}
    init_db()
    upserted = 0
    with get_session() as session:
        for person in people:
            pid = str(person["id"]).strip()
            if not pid:
                continue
            row = session.get(Personnel, pid)
            payload = {
                "fio": _normalize_fio(str(person.get("fio") or "")),
                "phone": str(person.get("phone") or "").strip(),
                "position": str(person.get("position") or "").strip(),
                "control_mode": str(person.get("control_mode") or "").strip(),
            }
            if not payload["fio"]:
                continue
            if row:
                for key, val in payload.items():
                    setattr(row, key, val)
            else:
                session.add(Personnel(id=pid, **payload))
            upserted += 1
    return {"upserted": upserted, "total": len(people)}


def import_personnel_xlsx_to_db(path) -> dict[str, Any]:
    from sk_reporter.personnel_xlsx import parse_personnel_rows

    people = parse_personnel_rows(path)
    if not people:
        raise ValueError("В файле нет строк сотрудников")
    result = upsert_people(people)
    result["source"] = "xlsx"
    return result
