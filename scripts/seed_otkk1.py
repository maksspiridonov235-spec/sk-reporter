#!/usr/bin/env python3
"""Залить эталон ОТКК-1 (6 пунктов) в PostgreSQL."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from sk_reporter.db.config import database_enabled
from sk_reporter.otkk_db import purge_empty_otkk_cards, seed_otkk1


def main() -> int:
    parser = argparse.ArgumentParser(description="Эталон ОТКК-1 в PostgreSQL")
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Перезаписать, если otkk-1 уже есть",
    )
    args = parser.parse_args()

    if not database_enabled():
        print("DATABASE_URL не задан", file=sys.stderr)
        return 1

    try:
        result = seed_otkk1(overwrite=args.overwrite)
    except Exception as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1

    if result.get("skipped"):
        print(f"Пропуск: {result['id']} уже в базе (используйте --overwrite)")
        return 0

    print(f"OK {result['id']}: {result.get('code')} — пунктов: {result.get('rows')}")
    purged = purge_empty_otkk_cards()
    if purged:
        print(f"  удалено пустых записей: {purged}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
