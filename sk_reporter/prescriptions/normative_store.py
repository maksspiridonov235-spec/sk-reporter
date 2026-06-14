"""
Локальная база нормативных документов для проверки предписаний.

Каталог: data/normative/ — manifest.yaml + текстовые файлы (см. README там).
Без Техэксперт и без интернет-поиска.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import yaml

from sk_reporter.paths import data_dir
from sk_reporter.prescriptions.techexpert_client import (
    _extract_points_excerpt,
    _short_doc_title,
    _title_from_reference,
    parse_normative_reference,
)

_NORMATIVE_DIR = data_dir() / "normative"
_MANIFEST_PATH = _NORMATIVE_DIR / "manifest.yaml"
_MAX_EXCERPT = 12_000


def normative_dir() -> Path:
    return _NORMATIVE_DIR


def manifest_path() -> Path:
    return _MANIFEST_PATH


def _load_manifest() -> list[dict[str, Any]]:
    if not _MANIFEST_PATH.is_file():
        return []
    raw = yaml.safe_load(_MANIFEST_PATH.read_text(encoding="utf-8")) or {}
    docs = raw.get("documents") or []
    return [d for d in docs if isinstance(d, dict)]


def normative_store_status() -> dict[str, Any]:
    docs = _load_manifest()
    texts_dir = _NORMATIVE_DIR / "texts"
    return {
        "manifest_exists": _MANIFEST_PATH.is_file(),
        "manifest_path": str(_MANIFEST_PATH),
        "documents_count": len(docs),
        "texts_dir_exists": texts_dir.is_dir(),
    }


def _norm_date_key(date: str) -> str:
    m = re.match(r"(\d{1,2})[./](\d{1,2})[./](\d{2,4})", (date or "").strip())
    if not m:
        return (date or "").strip()
    dd, mm, yy = m.groups()
    if len(yy) == 2:
        yy = "20" + yy
    return f"{int(dd):02d}.{int(mm):02d}.{yy}"


def _score_entry(entry: dict[str, Any], ref) -> int:
    score = 0
    number = str(entry.get("number") or "").strip()
    kind = str(entry.get("kind") or "").strip()
    issuer = str(entry.get("issuer") or "").strip()
    date = str(entry.get("date") or "").strip()

    if ref.number and number and ref.number == number:
        score += 10
    elif ref.number and number and ref.number in number:
        score += 4

    if ref.doc_kind and kind and ref.doc_kind.lower() == kind.lower():
        score += 3

    if ref.issuer and issuer and ref.issuer.lower()[:8] in issuer.lower():
        score += 4

    if ref.date and date and _norm_date_key(ref.date) == _norm_date_key(date):
        score += 3

    title = str(entry.get("title") or "")
    if ref.number and ref.number in title:
        score += 2

    return score


def _resolve_text_path(entry: dict[str, Any]) -> Path | None:
    rel = str(entry.get("file") or "").strip()
    if not rel:
        return None
    path = (_NORMATIVE_DIR / rel).resolve()
    root = _NORMATIVE_DIR.resolve()
    if not str(path).startswith(str(root)):
        return None
    return path if path.is_file() else None


def _read_document_text(path: Path) -> str:
    for enc in ("utf-8", "utf-8-sig", "cp1251"):
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def lookup_normative(normative_text: str) -> dict[str, Any]:
    """Поиск фрагмента НД в локальной базе data/normative/."""
    reference = parse_normative_reference(normative_text)
    base_ref = {
        "raw": reference.raw,
        "search_query": reference.search_query,
        "doc_kind": reference.doc_kind,
        "number": reference.number,
        "date": reference.date,
        "issuer": reference.issuer,
        "points": reference.points,
    }

    if not reference.raw:
        return {
            "ok": False,
            "reference": base_ref,
            "error": "Пустая ссылка на нормативный документ (B19)",
            "source": "local",
        }

    entries = _load_manifest()
    if not entries:
        return {
            "ok": False,
            "reference": base_ref,
            "error": (
                "Локальная база пуста — добавьте документы в data/normative/ "
                "(см. data/normative/README.md)"
            ),
            "source": "local",
        }

    ranked = sorted(
        (( _score_entry(e, reference), e) for e in entries),
        key=lambda x: x[0],
        reverse=True,
    )
    best_score, best_entry = ranked[0]
    if best_score < 6:
        hint = reference.number or reference.search_query[:60] or reference.raw[:60]
        return {
            "ok": False,
            "reference": base_ref,
            "error": (
                f"Документ не найден в локальной базе "
                f"(запись по B19: «{hint}»). Добавьте в manifest.yaml."
            ),
            "source": "local",
        }

    text_path = _resolve_text_path(best_entry)
    if not text_path:
        rel = best_entry.get("file") or "?"
        return {
            "ok": False,
            "reference": base_ref,
            "error": f"В manifest указан файл «{rel}», но на диске его нет",
            "source": "local",
        }

    plain = _read_document_text(text_path).strip()
    if len(plain) < 80:
        return {
            "ok": False,
            "reference": base_ref,
            "error": f"Файл {text_path.name} слишком короткий или пустой",
            "source": "local",
        }

    title = str(best_entry.get("title") or "")
    short = _short_doc_title(title, reference) or _title_from_reference(reference)
    excerpt = _extract_points_excerpt(plain, reference.points) or plain[:_MAX_EXCERPT]

    return {
        "ok": True,
        "reference": base_ref,
        "excerpt": excerpt,
        "doc_title": short,
        "list_title": title or short,
        "source_url": str(text_path.relative_to(_NORMATIVE_DIR)),
        "error": "",
        "auth_ok": True,
        "source": "local",
        "local_id": str(best_entry.get("id") or ""),
    }
