#!/usr/bin/env python3
"""Удалить пустые записи ОТКК (без content) из PostgreSQL."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from sk_reporter.db.config import database_enabled
from sk_reporter.otkk_db import db_status, purge_empty_otkk_cards


def main() -> int:
    if not database_enabled():
        print("DATABASE_URL не задан", file=sys.stderr)
        return 1
    deleted = purge_empty_otkk_cards()
    st = db_status()
    print(f"Удалено пустых записей: {deleted}. В базе карт: {st.get('count', 0)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
