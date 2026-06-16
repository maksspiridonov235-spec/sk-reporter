"""Каталог технологических карт (ТК) и сопоставление с видами работ."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

import yaml

from sk_reporter.engineer.doc_text import control_snippet_from_tk
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


def tk_text_for_id(tk_id: str) -> str:
    """Текст карты только из PostgreSQL."""
    try:
        from sk_reporter.otkk_parser import content_to_plain_text
        from sk_reporter.otkk_store import get_card

        card = get_card(tk_id, include_content=True)
        if card and card.get("content"):
            return content_to_plain_text(card["content"])
    except RuntimeError:
        pass
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
