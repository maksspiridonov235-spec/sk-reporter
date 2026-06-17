"""Импорт проекта с диска (data/projects/<id>/) → структура для PostgreSQL."""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from sk_reporter.engineer.vor_legacy import parse_vor_file
from sk_reporter.engineer.vor_parser import vor_to_dict
from sk_reporter.paths import project_dir
from sk_reporter.project_tl_parser import parse_tl_docx


def _find_vor_file(folder: Path) -> Path | None:
    for pattern in ("*ВОР*.doc", "*ВОР*.docx", "*вор*.doc", "*вор*.docx"):
        matches = sorted(folder.glob(pattern), key=lambda p: p.name.casefold())
        if matches:
            return matches[0]
    return None


def _find_tl_file(folder: Path) -> Path | None:
    for pattern in ("*ТЛ*.docx", "*ТЛ*.doc", "*тл*.docx"):
        matches = sorted(folder.glob(pattern), key=lambda p: p.name.casefold())
        if matches:
            return matches[0]
    return None


def _guess_titles(project_id: str, tl: dict[str, Any] | None) -> tuple[str, str]:
    title = project_id
    object_name = project_id
    if not tl:
        return title, object_name
    for table in tl.get("tables") or []:
        for row in table.get("rows") or []:
            for cell in row:
                if "Заказчик" in cell:
                    object_name = cell.split("–", 1)[-1].split("-", 1)[-1].strip()
                    return title, object_name
    return title, object_name


def build_project_content(folder: Path, *, project_id: str | None = None) -> dict[str, Any]:
    pid = project_id or folder.name
    vor_path = _find_vor_file(folder)
    tl_path = _find_tl_file(folder)
    if not vor_path and not tl_path:
        raise FileNotFoundError(f"В {folder} нет файлов ВОР или ТЛ")

    content: dict[str, Any] = {"imported_at": datetime.now(tz=timezone.utc).isoformat()}
    vor_file = ""
    tl_file = ""

    if vor_path:
        vor_doc = parse_vor_file(vor_path, pid)
        content["vor"] = vor_to_dict(vor_doc)
        vor_file = vor_path.name

    tl_data: dict[str, Any] | None = None
    if tl_path:
        if tl_path.suffix.lower() == ".docx":
            tl_data = parse_tl_docx(tl_path)
        else:
            tl_data = {"source": tl_path.name, "tables": [], "note": "ТЛ .doc — позже"}
        content["tl"] = tl_data
        tl_file = tl_path.name

    title, object_name = _guess_titles(pid, tl_data)
    return {
        "id": pid,
        "title": title,
        "object_name": object_name,
        "vor_file": vor_file,
        "tl_file": tl_file,
        "content": content,
    }


def import_project_folder(project_id: str) -> dict[str, Any]:
    folder = project_dir(project_id)
    if not folder.is_dir():
        raise FileNotFoundError(f"Нет папки проекта: {folder}")
    return build_project_content(folder, project_id=project_id)
