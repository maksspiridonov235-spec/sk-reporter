#!/usr/bin/env python3
"""Проверка доступа к Техэксперт и поиска нормативки (B19).

Перед запуском:
  export TE_EXPERT_LOGIN=...
  export TE_EXPERT_PASSWORD=...

Пример:
  TE_EXPERT_INTERNET_FALLBACK=0 python scripts/test_techexpert.py
"""

from __future__ import annotations

import json
import os
import sys

from sk_reporter.prescriptions.te_env import load_te_expert_env, te_expert_config_status
from sk_reporter.prescriptions.techexpert_client import (
    TechExpertClient,
    _build_search_queries,
    lookup_normative,
    parse_normative_reference,
)

load_te_expert_env()

SAMPLE = (
    "Приказ Ростехнадзора от 11.12.2020 N 519 "
    "(ред. от 14.01.2025), п. 44"
)


def main() -> int:
    text = sys.argv[1] if len(sys.argv) > 1 else SAMPLE
    ref = parse_normative_reference(text)
    print("B19:", text)
    print("parsed:", json.dumps(ref.__dict__, ensure_ascii=False, indent=2))
    print("queries:", _build_search_queries(ref))

    cfg = te_expert_config_status()
    print("config:", json.dumps(cfg, ensure_ascii=False, indent=2))
    if not os.environ.get("TE_EXPERT_LOGIN") or not os.environ.get("TE_EXPERT_PASSWORD"):
        print(
            "\nОШИБКА: задайте TE_EXPERT_LOGIN и TE_EXPERT_PASSWORD "
            "или создайте data/local/te_expert.env"
        )
        return 1

    client = TechExpertClient()
    auth_ok, auth_err = client._login_http()
    print("\n[1] login:", "OK" if auth_ok else f"FAIL — {auth_err}")
    if not auth_ok:
        return 2

    http = client._search_http(ref)
    print("\n[2] HTTP search (ifind API = строка поиска):")
    print("  ok:", http.ok)
    print("  error:", http.error or "—")
    print("  url:", http.source_url or "—")
    print("  title:", (http.doc_title or "")[:120])
    print("  excerpt_len:", len(http.excerpt or ""))
    if http.excerpt:
        print("  excerpt[:400]:", http.excerpt[:400])

    ui = client._search_browser(ref)
    print("\n[3] UI search (frame=left → клик → center):")
    print("  ok:", ui.ok)
    print("  error:", ui.error or "—")
    print("  url:", ui.source_url or "—")
    print("  excerpt_len:", len(ui.excerpt or ""))

    fallback = os.environ.get("TE_EXPERT_INTERNET_FALLBACK", "1")
    print(f"\n[4] lookup_normative (internet_fallback={fallback}):")
    r = lookup_normative(text)
    print("  ok:", r.get("ok"))
    print("  source:", r.get("source"))
    print("  te_fallback_error:", r.get("te_fallback_error") or "—")
    print("  error:", r.get("error") or "—")
    print("  url:", r.get("source_url") or "—")

    if r.get("source") == "techexpert" and r.get("ok"):
        print("\nИТОГ: Техэксперт работает.")
        return 0
    if r.get("source") == "internet" and r.get("ok"):
        print("\nИТОГ: использован интернет (см. te_fallback_error выше).")
        return 3
    print("\nИТОГ: нормативка не получена.")
    return 4


if __name__ == "__main__":
    raise SystemExit(main())
