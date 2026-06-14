#!/usr/bin/env python3
"""Проверка локальной базы normative (data/normative/)."""

from __future__ import annotations

import json
import sys

from sk_reporter.prescriptions.normative_store import (
    lookup_normative,
    normative_store_status,
)

SAMPLE = "Приказ Ростехнадзора от 11.12.2020 N 519, п. 44"


def main() -> int:
    text = sys.argv[1] if len(sys.argv) > 1 else SAMPLE
    print("status:", json.dumps(normative_store_status(), ensure_ascii=False, indent=2))
    print("\nB19:", text)
    r = lookup_normative(text)
    print("ok:", r.get("ok"))
    print("source:", r.get("source"))
    print("error:", r.get("error") or "—")
    print("doc_title:", r.get("doc_title") or "—")
    print("excerpt_len:", len(r.get("excerpt") or ""))
    if r.get("excerpt"):
        print("excerpt[:400]:", (r["excerpt"] or "")[:400])
    return 0 if r.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
