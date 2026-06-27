"""Добавление колонок в существующие таблицы без alembic."""

from __future__ import annotations

from sqlalchemy import create_engine, text
from sqlalchemy.exc import ProgrammingError

from sk_reporter.db.config import database_url


def _is_privilege_error(exc: BaseException) -> bool:
    msg = str(exc).lower()
    return "insufficientprivilege" in msg or "must be owner" in msg


def ensure_table_columns(
    table: str,
    *,
    add_columns: dict[str, str] | None = None,
    drop_not_null: list[str] | None = None,
) -> None:
    """Добавить отсутствующие колонки; ALTER только если реально нужен."""
    url = database_url()
    if not url:
        return

    add_columns = add_columns or {}
    drop_not_null = drop_not_null or []

    engine = create_engine(url, pool_pre_ping=True)
    with engine.begin() as conn:
        meta = {
            row[0]: row[1]
            for row in conn.execute(
                text(
                    "SELECT column_name, is_nullable FROM information_schema.columns "
                    "WHERE table_schema = 'public' AND table_name = :table"
                ),
                {"table": table},
            )
        }

        pending: list[str] = []
        for col, col_def in add_columns.items():
            if col not in meta:
                pending.append(f"ALTER TABLE {table} ADD COLUMN {col} {col_def}")
        for col in drop_not_null:
            if col in meta and meta[col] == "NO":
                pending.append(f"ALTER TABLE {table} ALTER COLUMN {col} DROP NOT NULL")

        if not pending:
            return

        try:
            for stmt in pending:
                conn.execute(text(stmt))
        except ProgrammingError as exc:
            if _is_privilege_error(exc):
                raise RuntimeError(
                    f"Нужна миграция БД (выполнить от владельца таблицы {table}):\n"
                    + "\n".join(pending)
                ) from exc
            raise
