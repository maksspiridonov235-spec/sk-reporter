"""Фасад проектов: PostgreSQL."""

from __future__ import annotations

from typing import Any

from sk_reporter.db.config import database_enabled


def _db():
    from sk_reporter import project_db

    return project_db


def engineer_project_map() -> dict[str, list[dict[str, str]]]:
    if not database_enabled():
        return {}
    return _db().engineer_project_map()


def get_project(project_id: str) -> dict[str, Any] | None:
    if not database_enabled():
        return None
    return _db().get_project(project_id)
