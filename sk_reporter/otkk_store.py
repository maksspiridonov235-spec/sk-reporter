"""Каталог ОТКК в PostgreSQL (RelaxDev, DATABASE_URL)."""

from __future__ import annotations

from typing import Any

from sk_reporter.db.config import database_enabled


def _require_database() -> None:
    if not database_enabled():
        raise RuntimeError(
            "DATABASE_URL не задан — каталог ОТКК хранится только в PostgreSQL"
        )


def load_cards() -> list[dict[str, Any]]:
    _require_database()
    from sk_reporter.otkk_db import load_cards_from_db

    return load_cards_from_db(with_content_only=True)


def get_card(card_id: str, *, include_content: bool = False) -> dict[str, Any] | None:
    _require_database()
    from sk_reporter.otkk_db import get_card_from_db

    card = get_card_from_db(card_id, include_content=include_content)
    if not card or not card.get("has_content"):
        return None
    return card
