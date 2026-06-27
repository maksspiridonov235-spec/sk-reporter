"""Сопоставление видов работ с картами ОТКК (PostgreSQL)."""

from __future__ import annotations

from pathlib import Path
from typing import Optional

import yaml

from sk_reporter.paths import project_dir


def load_work_tk_map(project_id: str) -> dict[str, str]:
    path = project_dir(project_id) / "work_tk_map.yaml"
    if not path.is_file():
        return {}
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    return {str(k): str(v) for k, v in (data.get("mappings") or {}).items()}


def resolve_tk_for_work(work_name: str, project_id: str) -> Optional[str]:
    mapping = load_work_tk_map(project_id)
    work_lower = work_name.lower()
    best_key = ""
    best_val: Optional[str] = None
    for key, val in mapping.items():
        if key.lower() in work_lower and len(key) > len(best_key):
            best_key = key
            best_val = val
    return best_val


def tk_text_for_id(tk_id: str) -> str:
    """Текст карты только из PostgreSQL."""
    try:
        from sk_reporter.otkk_parser import content_to_plain_text
        from sk_reporter.otkk_store import get_card

        card = get_card(tk_id, include_content=True)
        if card and card.get("content"):
            return content_to_plain_text(card["content"])
    except RuntimeError:
        pass
    return ""
