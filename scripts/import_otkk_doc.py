#!/usr/bin/env python3
"""Импорт одной ОТКК из .doc/.docx в PostgreSQL (структура + файл на диск)."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from sk_reporter.db.config import database_enabled
from sk_reporter.otkk_db import import_document_to_db


def main() -> int:
    parser = argparse.ArgumentParser(description="Импорт ОТКК в PostgreSQL")
    parser.add_argument("path", type=Path, help="Путь к .doc или .docx")
    parser.add_argument(
        "--no-copy",
        action="store_true",
        help="Не копировать файл в data/tk/ (только БД)",
    )
    args = parser.parse_args()

    if not database_enabled():
        print("DATABASE_URL не задан", file=sys.stderr)
        return 1

    try:
        result = import_document_to_db(args.path, copy_to_tk_dir=not args.no_copy)
    except Exception as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1

    title = result.get("title") or ""
    short = title if len(title) <= 60 else title[:60] + "…"
    print(f"OK {result['id']}: {result.get('code')} — {short}")
    print(f"  файл: {result.get('file')}, строк таблицы: {result.get('rows')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
