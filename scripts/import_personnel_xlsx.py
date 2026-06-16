#!/usr/bin/env python3
"""Импорт Excel со списком сотрудников в PostgreSQL (нужен DATABASE_URL)."""

from __future__ import annotations

import sys

from sk_reporter.db.config import database_enabled
from sk_reporter.personnel_db import db_status, import_personnel_xlsx_to_db
from sk_reporter.paths import personnel_dir


def main() -> int:
    if not database_enabled():
        print("DATABASE_URL не задан", file=sys.stderr)
        return 1
    xlsx = personnel_dir() / "Справочник персонала.xlsx"
    if not xlsx.is_file():
        print(f"Файл не найден: {xlsx}", file=sys.stderr)
        return 1
    try:
        result = import_personnel_xlsx_to_db(xlsx)
    except Exception as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1
    st = db_status()
    print(
        f"Импортировано: {result.get('upserted', 0)} из {result.get('total', 0)} "
        f"(источник: {result.get('source')}). В базе: {st.get('count', 0)}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
