"""Каталог технологических карт (ТК) и сопоставление с видами работ."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

import yaml

from sk_reporter.engineer.doc_text import control_snippet_from_tk, extract_doc_text
from sk_reporter.paths import tk_dir

_OTKK_RE = re.compile(r"ОТКК[-\s]?(\d+)", re.I)


def list_tk_files(root: Optional[Path] = None) -> list[dict]:
    root = root or tk_dir()
    items: list[dict] = []
    for path in sorted(root.iterdir()):
        if path.suffix.lower() not in {".doc", ".docx"}:
            continue
        m = _OTKK_RE.search(path.name)
        otkk_id = f"otkk-{m.group(1)}" if m else path.stem
        items.append({"id": otkk_id, "file": path.name})
    return items


def write_manifest(root: Optional[Path] = None) -> Path:
    root = root or tk_dir()
    manifest = {"cards": list_tk_files(root)}
    out = root / "manifest.yaml"
    out.write_text(yaml.safe_dump(manifest, allow_unicode=True, sort_keys=False), encoding="utf-8")
    return out


def load_work_tk_map(project_dir: Path) -> dict[str, str]:
    path = project_dir / "work_tk_map.yaml"
    if not path.is_file():
        return {}
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    return {str(k): str(v) for k, v in (data.get("mappings") or {}).items()}


def resolve_tk_for_work(work_name: str, project_dir: Path) -> Optional[str]:
    mapping = load_work_tk_map(project_dir)
    work_lower = work_name.lower()
    best_key = ""
    best_val: Optional[str] = None
    for key, val in mapping.items():
        if key.lower() in work_lower and len(key) > len(best_key):
            best_key = key
            best_val = val
    return best_val


def resolve_tk_path(tk_id: str, root: Optional[Path] = None) -> Optional[Path]:
    root = root or tk_dir()
    try:
        from sk_reporter.otkk_store import card_file_path, get_card

        card = get_card(tk_id)
        if card:
            p = card_file_path(card, root)
            if p.is_file():
                return p
    except RuntimeError:
        pass

    manifest_path = root / "manifest.yaml"
    if manifest_path.is_file():
        data = yaml.safe_load(manifest_path.read_text(encoding="utf-8")) or {}
        for card in data.get("cards") or []:
            if card.get("id") == tk_id:
                p = root / Path(str(card.get("file") or "")).name
                if p.is_file():
                    return p
    tk_id_norm = tk_id.lower().replace("otkk-", "ОТКК-")
    for path in root.iterdir():
        if path.suffix.lower() not in {".doc", ".docx"}:
            continue
        if tk_id_norm in path.name.upper() or tk_id.lower() in path.name.lower():
            return path
    return None


def extract_tk_text(path: Path) -> str:
    return extract_doc_text(path)


def tk_text_for_id(tk_id: str, root: Optional[Path] = None) -> str:
    """Текст карты: из PostgreSQL (структура), иначе с диска."""
    try:
        from sk_reporter.otkk_parser import content_to_plain_text
        from sk_reporter.otkk_store import get_card

        card = get_card(tk_id, include_content=True)
        if card and card.get("content"):
            return content_to_plain_text(card["content"])
    except RuntimeError:
        pass
    tk_path = resolve_tk_path(tk_id, root)
    if not tk_path:
        return ""
    try:
        return extract_tk_text(tk_path)
    except Exception:
        return ""


def snippet_for_work(work_name: str, project_dir: Path, max_chars: int = 900) -> str:
    tk_id = resolve_tk_for_work(work_name, project_dir)
    if not tk_id:
        return ""
    try:
        text = tk_text_for_id(tk_id)
        if not text:
            return ""
        return control_snippet_from_tk(text, max_chars=max_chars)
    except Exception:
        return ""
