"""Пути к данным репозитория (шаблоны, корень проекта)."""

from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parent.parent


def repo_root() -> Path:
    return _REPO_ROOT


def data_dir() -> Path:
    return _REPO_ROOT / "data"


def templates_dir() -> Path:
    return data_dir() / "templates"


def projects_dir() -> Path:
    return data_dir() / "projects"


def project_dir(project_id: str) -> Path:
    return projects_dir() / project_id


def tk_dir() -> Path:
    return data_dir() / "tk"


def deployment_dir() -> Path:
    return data_dir() / "deployment"


def default_rasstanovka_template_path() -> Path:
    return deployment_dir() / "rasstanovka_template.xlsm"


def default_pril7_template_path() -> Path:
    return deployment_dir() / "pril7_template.xlsm"

