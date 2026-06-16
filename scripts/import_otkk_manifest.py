#!/usr/bin/env python3
"""Импорт data/tk/manifest.yaml в PostgreSQL (нужен DATABASE_URL)."""

from __future__ import annotations

import sys

from sk_reporter.db.config import database_enabled
from sk_reporter.otkk_db import db_status, import_manifest_to_db
from sk_reporter.paths import tk_dir


def main() -> int:
    if not database_enabled():
        print("DATABASE_URL не задан", file=sys.stderr)
        return 1
    manifest = tk_dir() / "manifest.yaml"
    if not manifest.is_file():
        print(f"Файл не найден: {manifest}", file=sys.stderr)
        return 1
    try:
        result = import_manifest_to_db(manifest)
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
