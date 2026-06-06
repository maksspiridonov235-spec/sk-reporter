"""Профиль инженера (yaml + env SK_ENGINEER_PROFILE)."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import yaml

from sk_reporter.paths import engineer_profiles_dir, repo_root


def load_profile(profile_id: str | None = None) -> dict[str, Any]:
    pid = profile_id or os.environ.get("SK_ENGINEER_PROFILE", "").strip()
    if not pid:
        raise ValueError("Не задан профиль: SK_ENGINEER_PROFILE или параметр profile_id")

    path = engineer_profiles_dir() / f"{pid}.yaml"
    if not path.is_file():
        raise FileNotFoundError(f"Профиль не найден: {path}")

    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    data["id"] = data.get("id") or pid
    return data


def resolve_report_template(profile: dict[str, Any]) -> Path | None:
    raw = profile.get("report_template")
    if not raw:
        return None
    p = Path(raw)
    if not p.is_absolute():
        p = repo_root() / p
    return p if p.is_file() else None
