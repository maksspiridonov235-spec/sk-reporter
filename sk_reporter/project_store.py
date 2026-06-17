"""Проекты — заглушка до переноса в PostgreSQL (yaml-слой отключён)."""

from __future__ import annotations

from typing import Any


def engineer_project_map() -> dict[str, list[dict[str, str]]]:
    return {}


def get_project(project_id: str) -> dict[str, Any] | None:
    return None


def list_projects_rich() -> list[dict[str, Any]]:
    return []


def set_project_engineers(project_id: str, engineer_ids: list[str]) -> dict[str, Any]:
    raise RuntimeError("Назначения проектов переносятся в PostgreSQL — yaml отключён")
