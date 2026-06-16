"""Подключение к PostgreSQL (RelaxDev: переменная DATABASE_URL)."""

from __future__ import annotations

import os


def database_url() -> str | None:
    raw = (os.environ.get("DATABASE_URL") or "").strip()
    return raw or None


def database_enabled() -> bool:
    return database_url() is not None
