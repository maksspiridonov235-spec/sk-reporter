"""Пути к шаблонам расстановки (в репозитории и загруженные пользователем)."""

from __future__ import annotations

from pathlib import Path

from sk_reporter.paths import default_rasstanovka_template_path

# Загруженный в temp — переопределяет шаблон из репозитория
_UPLOADED_TEMPLATE_NAME = "template.xlsm"


def uploaded_template_path(work_deployment_dir: Path) -> Path:
    return work_deployment_dir / _UPLOADED_TEMPLATE_NAME


def resolve_rasstanovka_template(work_deployment_dir: Path) -> tuple[Path, str]:
    """(путь, источник: 'upload' | 'bundled')."""
    uploaded = uploaded_template_path(work_deployment_dir)
    if uploaded.is_file():
        return uploaded, "upload"
    bundled = default_rasstanovka_template_path()
    if bundled.is_file():
        return bundled, "bundled"
    raise FileNotFoundError("Шаблон расстановки не найден")


def template_status(work_deployment_dir: Path) -> dict:
    uploaded = uploaded_template_path(work_deployment_dir)
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
