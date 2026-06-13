"""Локальные TE_EXPERT_* из data/local/te_expert.env (файл не в git)."""

from __future__ import annotations

import os
from pathlib import Path

from sk_reporter.paths import repo_root

_LOADED = False
_ENV_PATH = repo_root() / "data" / "local" / "te_expert.env"


def te_expert_env_path() -> Path:
    return _ENV_PATH


def _should_load_env_value(key: str, value: str) -> bool:
    if not key or not value:
        return False
    current = os.environ.get(key)
    if current is None:
        return True
    return not str(current).strip()


def load_te_expert_env(force: bool = False) -> bool:
    """Подставляет TE_EXPERT_* из локального файла (перезаписывает пустые env)."""
    global _LOADED
    if _LOADED and not force:
        return _ENV_PATH.is_file()

    _LOADED = True
    if not _ENV_PATH.is_file():
        return False

    for raw_line in _ENV_PATH.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.lower().startswith("export "):
            line = line[7:].strip()
        if "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip().lstrip("\ufeff")
        value = value.strip().strip('"').strip("'")
        if _should_load_env_value(key, value):
            os.environ[key] = value
    return True


def te_expert_config_status() -> dict[str, object]:
    load_te_expert_env()
    login = os.environ.get("TE_EXPERT_LOGIN", "").strip()
    password = os.environ.get("TE_EXPERT_PASSWORD", "").strip()
    return {
        "env_file": str(_ENV_PATH),
        "env_file_exists": _ENV_PATH.is_file(),
        "login_set": bool(login),
        "password_set": bool(password),
        "configured": bool(login and password),
        "base_url": os.environ.get("TE_EXPERT_BASE_URL", "").strip()
        or "http://248960.te-cloud.ru",
        "internet_fallback": os.environ.get("TE_EXPERT_INTERNET_FALLBACK", "1"),
    }
