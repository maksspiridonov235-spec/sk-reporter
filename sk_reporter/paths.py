"""Пути к данным репозитория (шаблоны, корень проекта)."""

from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parent.parent

def repo_root() -> Path:
    return _REPO_ROOT


def templates_dir() -> Path:
    return _REPO_ROOT / "data" / "templates"


def output_dir() -> Path:
    """Папка исправленных отчётов («Проверить и исправить»)."""
    d = _REPO_ROOT / "output"
    d.mkdir(exist_ok=True)
    return d
