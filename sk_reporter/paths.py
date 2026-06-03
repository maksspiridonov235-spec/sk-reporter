"""Пути к данным репозитория (шаблоны, корень проекта)."""

from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parent.parent

def repo_root() -> Path:
    return _REPO_ROOT


def templates_dir() -> Path:
    return _REPO_ROOT / "data" / "templates"
