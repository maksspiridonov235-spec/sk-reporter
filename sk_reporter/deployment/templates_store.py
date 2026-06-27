"""Пути к шаблонам расстановки и Прил.7 (в репозитории и загруженные пользователем)."""

from __future__ import annotations

from pathlib import Path

from sk_reporter.paths import default_pril7_template_path, default_rasstanovka_template_path

# Загруженные в temp — переопределяют шаблоны из репозитория
_UPLOADED_RASSTANOVKA_NAME = "template.xlsm"
_UPLOADED_PRIL7_NAME = "pril7.xlsm"


def uploaded_rasstanovka_path(work_deployment_dir: Path) -> Path:
    return work_deployment_dir / _UPLOADED_RASSTANOVKA_NAME


def uploaded_pril7_path(work_deployment_dir: Path) -> Path:
    return work_deployment_dir / _UPLOADED_PRIL7_NAME


def resolve_rasstanovka_template(work_deployment_dir: Path) -> tuple[Path, str]:
    """(путь, источник: 'upload' | 'bundled')."""
    uploaded = uploaded_rasstanovka_path(work_deployment_dir)
    if uploaded.is_file():
        return uploaded, "upload"
    bundled = default_rasstanovka_template_path()
    if bundled.is_file():
        return bundled, "bundled"
    raise FileNotFoundError("Шаблон расстановки не найден")


def resolve_pril7_template(work_deployment_dir: Path) -> tuple[Path, str]:
    """(путь, источник: 'upload' | 'bundled')."""
    uploaded = uploaded_pril7_path(work_deployment_dir)
    if uploaded.is_file():
        return uploaded, "upload"
    bundled = default_pril7_template_path()
    if bundled.is_file():
        return bundled, "bundled"
    raise FileNotFoundError("Шаблон Приложения 7 не найден")


def template_status(work_deployment_dir: Path) -> dict:
    uploaded = uploaded_rasstanovka_path(work_deployment_dir)
    bundled = default_rasstanovka_template_path()
    if uploaded.is_file():
        return {
            "available": True,
            "source": "upload",
            "name": uploaded.name,
            "bundled_name": bundled.name if bundled.is_file() else None,
        }
    if bundled.is_file():
        return {
            "available": True,
            "source": "bundled",
            "name": bundled.name,
            "bundled_name": bundled.name,
        }
    return {"available": False, "source": None, "name": None, "bundled_name": None}


def pril7_status(work_deployment_dir: Path) -> dict:
    uploaded = uploaded_pril7_path(work_deployment_dir)
    bundled = default_pril7_template_path()
    if uploaded.is_file():
        return {
            "available": True,
            "source": "upload",
            "name": uploaded.name,
            "bundled_name": bundled.name if bundled.is_file() else None,
        }
    if bundled.is_file():
        return {
            "available": True,
            "source": "bundled",
            "name": bundled.name,
            "bundled_name": bundled.name,
        }
    return {"available": False, "source": None, "name": None, "bundled_name": None}
