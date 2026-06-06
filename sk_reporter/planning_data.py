"""Список файлов и метаданных для раздела «Планирование»."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml

from sk_reporter.paths import luvr_dir, personnel_dir, projects_dir, repo_root, tk_dir

_SECTIONS = frozenset({"projects", "luvr", "personnel", "otkk"})


def _file_row(path: Path) -> dict[str, Any]:
    st = path.stat()
    root = repo_root()
    try:
        rel = str(path.relative_to(root))
    except ValueError:
        rel = str(path)
    return {
        "name": path.name,
        "rel": rel,
        "size_kb": round(st.st_size / 1024, 1),
        "suffix": path.suffix.lower(),
    }


def _list_files(folder: Path, pattern: str = "*") -> list[dict[str, Any]]:
    if not folder.is_dir():
        return []
    rows = []
    for p in sorted(folder.glob(pattern)):
        if p.name.startswith("."):
            continue
        if p.is_file():
            rows.append(_file_row(p))
    return rows


def list_projects() -> list[dict[str, Any]]:
    out = []
    root = projects_dir()
    if not root.is_dir():
        return out
    for proj in sorted(root.iterdir()):
        if not proj.is_dir() or proj.name.startswith("."):
            continue
        meta: dict[str, Any] = {"id": proj.name, "title": proj.name}
        meta_path = proj / "project.yaml"
        if meta_path.is_file():
            meta.update(yaml.safe_load(meta_path.read_text(encoding="utf-8")) or {})
        files = []
        for p in sorted(proj.iterdir()):
            if p.is_file() and not p.name.startswith("."):
                files.append(_file_row(p))
        out.append(
            {
                "id": meta.get("id") or proj.name,
                "title": meta.get("title") or proj.name,
                "vor_docx": meta.get("vor_docx"),
                "has_vor_cache": (proj / "vor.json").is_file(),
                "files": files,
                "path": str(proj.relative_to(repo_root())),
            }
        )
    return out


def list_luvr() -> dict[str, Any]:
    folder = luvr_dir()
    return {
        "folder": str(folder.relative_to(repo_root())),
        "files": _list_files(folder),
    }


def list_personnel() -> dict[str, Any]:
    folder = personnel_dir()
    people_count = 0
    yaml_path = folder / "personnel.yaml"
    if yaml_path.is_file():
        data = yaml.safe_load(yaml_path.read_text(encoding="utf-8")) or {}
        people_count = len(data.get("people") or [])
    return {
        "folder": str(folder.relative_to(repo_root())),
        "people_count": people_count,
        "files": _list_files(folder),
    }


def list_otkk() -> dict[str, Any]:
    folder = tk_dir()
    cards = []
    manifest = folder / "manifest.yaml"
    if manifest.is_file():
        data = yaml.safe_load(manifest.read_text(encoding="utf-8")) or {}
        for card in data.get("cards") or []:
            fname = card.get("file") or ""
            fp = folder / fname
            cards.append(
                {
                    "id": card.get("id"),
                    "file": fname,
                    "present": fp.is_file(),
                    "size_kb": round(fp.stat().st_size / 1024, 1) if fp.is_file() else None,
                }
            )
    else:
        for p in sorted(folder.iterdir()):
            if p.suffix.lower() in {".doc", ".docx"} and not p.name.startswith("."):
                cards.append(
                    {
                        "id": p.stem[:20],
                        "file": p.name,
                        "present": True,
                        "size_kb": round(p.stat().st_size / 1024, 1),
                    }
                )
    return {
        "folder": str(folder.relative_to(repo_root())),
        "count": len(cards),
        "cards": cards,
    }


def planning_section(section: str) -> dict[str, Any]:
    if section not in _SECTIONS:
        raise KeyError(section)
    if section == "projects":
        return {"section": section, "items": list_projects()}
    if section == "luvr":
        return {"section": section, **list_luvr()}
    if section == "personnel":
        return {"section": section, **list_personnel()}
    return {"section": section, **list_otkk()}
