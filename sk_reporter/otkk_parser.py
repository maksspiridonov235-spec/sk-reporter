"""Плоский текст из JSON-структуры карточки ОТКК."""

from __future__ import annotations

from typing import Any


def content_to_plain_text(content: dict[str, Any]) -> str:
    """Плоский текст из структуры (для сниппетов в отчёте инженера)."""
    parts: list[str] = []
    if content.get("code"):
        parts.append(str(content["code"]))
    if content.get("title"):
        parts.append(str(content["title"]))
    for row in content.get("rows") or []:
        label = row.get("label") or ""
        value = row.get("value") or ""
        if label:
            parts.append(f"{label}: {value}")
        body = row.get("body") or {}
        for p in body.get("paragraphs") or []:
            parts.append(p)
        for b in body.get("bullets") or []:
            parts.append(f"- {b}")
    sig = content.get("signature") or {}
    if sig.get("text"):
        parts.append(sig["text"])
    return "\n".join(p for p in parts if p)
