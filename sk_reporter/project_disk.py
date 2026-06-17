"""Проекты на диске: data/projects/<id>/ — список для UI планирования."""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote

from sk_reporter.paths import project_dir, projects_dir


def _file_mtime_iso(path: Path) -> str:
    ts = path.stat().st_mtime
    return datetime.fromtimestamp(ts, tz=timezone.utc).astimezone().isoformat(timespec="seconds")


def _file_row(project_id: str, path: Path) -> dict[str, Any]:
    suffix = path.suffix.lower().lstrip(".")
    size = path.stat().st_size
    return {
        "name": path.name,
        "suffix": suffix or "—",
        "size_kb": round(size / 1024, 1),
        "size_bytes": size,
        "modified": _file_mtime_iso(path),
        "download_url": (
            f"/api/planning/projects/{quote(project_id, safe='')}"
            f"/files/{quote(path.name, safe='')}"
        ),
    }


def _list_files_in_project(project_id: str) -> list[dict[str, Any]]:
    root = project_dir(project_id)
    if not root.is_dir():
        return []
    files = [
        p
        for p in sorted(root.iterdir(), key=lambda x: x.name.casefold())
        if p.is_file() and not p.name.startswith(".")
    ]
    return [_file_row(project_id, p) for p in files]


def list_disk_projects() -> list[dict[str, Any]]:
    root = projects_dir()
    if not root.is_dir():
        return []
    out: list[dict[str, Any]] = []
    for entry in sorted(root.iterdir(), key=lambda x: x.name.casefold()):
        if not entry.is_dir() or entry.name.startswith("."):
            continue
        pid = entry.name
        files = _list_files_in_project(pid)
        out.append(
            {
                "id": pid,
                "folder": pid,
                "files_count": len(files),
                "files": files,
            }
        )
    return out


def get_disk_project(project_id: str) -> dict[str, Any] | None:
    pid = str(project_id).strip()
    if not pid:
        return None
    root = project_dir(pid)
    if not root.is_dir():
        return None
    files = _list_files_in_project(pid)
    return {
        "id": pid,
        "folder": pid,
        "files_count": len(files),
        "files": files,
    }


def resolve_project_file(project_id: str, filename: str) -> Path:
    pid = str(project_id).strip()
    name = Path(filename).name
    if not pid or not name or name != filename:
        raise FileNotFoundError("Недопустимое имя файла")
    base = project_dir(pid).resolve()
    if not base.is_dir():
        raise FileNotFoundError(f"Проект не найден: {pid}")
    target = (base / name).resolve()
    try:
        target.relative_to(base)
    except ValueError as exc:
        raise PermissionError("Путь вне каталога проекта") from exc
    if not target.is_file():
        raise FileNotFoundError(f"Файл не найден: {name}")
    return target


def disk_status() -> dict[str, Any]:
    root = projects_dir()
    projects = list_disk_projects()
    return {
        "root": str(root),
        "exists": root.is_dir(),
        "count": len(projects),
        "ok": root.is_dir(),
    }
