#!/usr/bin/env python3
"""Импорт «Ядро.xlsx» (листы Подрядчики + Объекты) в PostgreSQL."""

from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from sk_reporter.core_import import import_core_xlsx_to_db  # noqa: E402


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python scripts/import_core_xlsx.py /path/to/Ядро.xlsx", file=sys.stderr)
        return 2
    path = Path(sys.argv[1]).expanduser()
    result = import_core_xlsx_to_db(path)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
