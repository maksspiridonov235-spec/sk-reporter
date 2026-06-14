#!/usr/bin/env python3
"""Проверка Ollama (local или cloud через OLLAMA_API_KEY)."""

from __future__ import annotations

import sys

from sk_reporter.llm_client import default_model, llm_chat, llm_status, ping_llm


def main() -> int:
    st = llm_status()
    print(f"mode={st['mode']} host={st['host']} model={st['model']} api_key={'yes' if st['api_key_set'] else 'no'}")
    ok, err = ping_llm()
    if not ok:
        print(f"ping FAIL: {err}")
        return 1
    print("ping OK")
    resp = llm_chat(
        messages=[{"role": "user", "content": "Ответь одним словом: ок"}],
        options={"num_predict": 10},
    )
    text = (resp.get("message") or {}).get("content") or ""
    print(f"chat sample: {text.strip()[:80]!r}")
    print(f"model={default_model()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
