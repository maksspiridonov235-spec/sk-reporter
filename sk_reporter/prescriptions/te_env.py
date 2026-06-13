"""Локальные TE_EXPERT_* из data/local/te_expert.env (файл не в git)."""

from __future__ import annotations

import os
from pathlib import Path

from sk_reporter.paths import repo_root

_ENV_PATH = repo_root() / "data" / "local" / "te_expert.env"
_EXAMPLE_PATH = repo_root() / "data" / "local" / "te_expert.env.example"
_TE_PREFIX = "TE_EXPERT_"
_PLACEHOLDER_LOGINS = {
    "ваш_логин",
    "your_login",
    "login",
    "логин",
}
_PLACEHOLDER_PASSWORDS = {
    "ваш_пароль",
    "your_password",
    "password",
    "пароль",
}


def te_expert_env_path() -> Path:
    return _ENV_PATH


def te_expert_env_example_path() -> Path:
    return _EXAMPLE_PATH


def _normalize_key(key: str) -> str:
    return key.strip().lstrip("\ufeff")


def _normalize_value(value: str) -> str:
    value = value.strip().strip('"').strip("'")
    if " #" in value:
        value = value.split(" #", 1)[0].rstrip()
    return value


def _is_placeholder(key: str, value: str) -> bool:
    val = value.strip().lower()
    if not val:
        return True
    if key == "TE_EXPERT_LOGIN":
        return val in _PLACEHOLDER_LOGINS
    if key == "TE_EXPERT_PASSWORD":
        return val in _PLACEHOLDER_PASSWORDS
    return False


def _read_env_file(path: Path) -> dict[str, str]:
    if not path.is_file():
        return {}

    raw = path.read_bytes()
    text = ""
    for encoding in ("utf-8-sig", "utf-8", "cp1251"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if not text:
        return {}

    out: dict[str, str] = {}
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.lower().startswith("export "):
            line = line[7:].strip()
        if "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = _normalize_key(key)
        value = _normalize_value(value)
        if not key.startswith(_TE_PREFIX) or not value:
            continue
        if _is_placeholder(key, value):
            continue
        out[key] = value
    return out


def parse_te_expert_env_file(path: Path | None = None) -> dict[str, str]:
    return _read_env_file(path or _ENV_PATH)


def load_te_expert_env(force: bool = False) -> bool:
    """Подставляет TE_EXPERT_* из data/local/te_expert.env (файл важнее env)."""
    if not _ENV_PATH.is_file():
        return False

    loaded_any = False
    for key, value in parse_te_expert_env_file().items():
        current = os.environ.get(key)
        if force or current != value:
            os.environ[key] = value
            loaded_any = True
    return loaded_any or bool(parse_te_expert_env_file())


def te_expert_config_status() -> dict[str, object]:
    file_values = parse_te_expert_env_file()
    load_te_expert_env()

    login = file_values.get("TE_EXPERT_LOGIN") or os.environ.get("TE_EXPERT_LOGIN", "")
    password = file_values.get("TE_EXPERT_PASSWORD") or os.environ.get("TE_EXPERT_PASSWORD", "")
    login = login.strip()
    password = password.strip()

    example_exists = _EXAMPLE_PATH.is_file()
    example_values = parse_te_expert_env_file(_EXAMPLE_PATH) if example_exists else {}
    example_has_real_login = bool(example_values.get("TE_EXPERT_LOGIN"))
    edited_example_only = (
        not _ENV_PATH.is_file()
        and example_has_real_login
        and bool(example_values.get("TE_EXPERT_PASSWORD"))
    )

    return {
        "env_file": str(_ENV_PATH),
        "env_file_exists": _ENV_PATH.is_file(),
        "example_file": str(_EXAMPLE_PATH),
        "example_file_exists": example_exists,
        "edited_example_only": edited_example_only,
        "login_set": bool(login),
        "password_set": bool(password),
        "configured": bool(login and password),
        "base_url": (
            file_values.get("TE_EXPERT_BASE_URL")
            or os.environ.get("TE_EXPERT_BASE_URL", "").strip()
            or "http://248960.te-cloud.ru"
        ),
        "internet_fallback": (
            file_values.get("TE_EXPERT_INTERNET_FALLBACK")
            or os.environ.get("TE_EXPERT_INTERNET_FALLBACK", "1")
        ),
    }
