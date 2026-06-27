"""Справочник должностей (описания для Прил.7 / расстановки) в PostgreSQL."""

from __future__ import annotations

import json
from typing import Any

from sk_reporter.db.config import database_enabled
from sk_reporter.db.models import Position
from sk_reporter.db.session import get_session, init_db
from sk_reporter.paths import data_dir


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
            count = session.query(Position).count()
        return {"enabled": True, "configured": True, "count": count, "ok": True}
    except Exception as exc:
        return {"enabled": True, "configured": True, "count": 0, "ok": False, "error": str(exc)}


def list_positions() -> list[dict[str, Any]]:
    init_db()
    with get_session() as session:
        rows = session.query(Position).order_by(Position.sort_order, Position.title).all()
        return [
            {"title": r.title, "description": r.description or "", "sort_order": r.sort_order}
            for r in rows
        ]


def position_sort_map() -> dict[str, int]:
    return {r["title"]: r["sort_order"] for r in list_positions()}


def position_descriptions_map() -> dict[str, str]:
    return {r["title"]: r["description"] for r in list_positions() if r["title"]}


def _json_seed_path():
    return data_dir() / "planning" / "position_descriptions.json"


def seed_positions_from_json(*, overwrite: bool = False) -> dict[str, Any]:
    """Залить должности из data/planning/position_descriptions.json."""
    path = _json_seed_path()
    if not path.is_file():
        return {"seeded": False, "reason": f"нет файла {path.name}"}
    if not database_enabled():
        return {"seeded": False, "reason": "DATABASE_URL не задан"}

    rows = json.loads(path.read_text(encoding="utf-8"))
    init_db()
    seeded = 0
    with get_session() as session:
        existing = session.query(Position).count()
        if existing and not overwrite:
            return {"seeded": False, "reason": "таблица positions уже заполнена", "count": existing}
        if overwrite and existing:
            session.query(Position).delete()
        for i, row in enumerate(rows):
            title = str(row.get("dolzhnost") or "").strip()
            if not title:
                continue
            session.merge(
                Position(
                    title=title,
                    description=str(row.get("opisanie") or "").strip(),
                    sort_order=i,
                )
            )
            seeded += 1
    return {"seeded": True, "count": seeded}
