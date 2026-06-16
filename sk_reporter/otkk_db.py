"""Каталог ОТКК (технологические карты) в PostgreSQL."""

from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any

from sqlalchemy import text
from sk_reporter.db.config import database_enabled, database_url
from sk_reporter.db.models import OtkkCard
from sk_reporter.db.session import get_session, init_db
from sk_reporter.otkk_text import sanitize_otkk_rows, strip_kodeks_fields
from sk_reporter.paths import tk_dir


def _ensure_otkk_schema() -> None:
    """Добавить колонки content/code/title в существующую таблицу (без alembic)."""
    from sqlalchemy import create_engine

    url = database_url()
    if not url:
        return
    engine = create_engine(url, pool_pre_ping=True)
    stmts = [
        "ALTER TABLE otkk_cards ADD COLUMN IF NOT EXISTS code VARCHAR(64) NOT NULL DEFAULT ''",
        "ALTER TABLE otkk_cards ADD COLUMN IF NOT EXISTS title TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE otkk_cards ADD COLUMN IF NOT EXISTS content JSONB",
    ]
    with engine.begin() as conn:
        for stmt in stmts:
            conn.execute(text(stmt))


def db_status() -> dict[str, Any]:
    if not database_enabled():
        return {
            "enabled": False,
            "configured": False,
            "count": 0,
            "with_content": 0,
            "ok": False,
            "error": "DATABASE_URL не задан",
        }
    try:
        init_db()
        _ensure_otkk_schema()
        with get_session() as session:
            with_content = session.query(OtkkCard).filter(OtkkCard.content.isnot(None)).count()
        return {
            "enabled": True,
            "configured": True,
            "count": with_content,
            "with_content": with_content,
            "ok": True,
        }
    except Exception as exc:
        return {"enabled": True, "configured": True, "count": 0, "with_content": 0, "ok": False, "error": str(exc)}


def _sanitize_content(content: dict[str, Any]) -> dict[str, Any]:
    """Старые записи в БД могли залиться с HYPERLINK — чистим при отдаче в UI."""
    if not content:
        return content
    out = dict(content)
    rows = out.get("rows")
    if rows:
        out["rows"] = sanitize_otkk_rows(rows)
    plain = out.get("plain_text")
    if plain:
        out["plain_text"] = strip_kodeks_fields(str(plain))
    return out


def _row_to_dict(row: OtkkCard, *, include_content: bool = False) -> dict[str, Any]:
    out: dict[str, Any] = {
        "id": row.id,
        "file": row.file_name,
        "code": row.code or "",
        "title": row.title or "",
        "has_content": row.content is not None,
    }
    if include_content and row.content is not None:
        out["content"] = _sanitize_content(row.content)
    return out


def load_cards_from_db(*, with_content_only: bool = False) -> list[dict[str, Any]]:
    init_db()
    _ensure_otkk_schema()
    with get_session() as session:
        q = session.query(OtkkCard).order_by(OtkkCard.id)
        if with_content_only:
            q = q.filter(OtkkCard.content.isnot(None))
        rows = q.all()
        return [_row_to_dict(r) for r in rows]


def purge_empty_otkk_cards() -> int:
    """Удалить строки каталога без content (остатки manifest/скана)."""
    init_db()
    _ensure_otkk_schema()
    with get_session() as session:
        return session.query(OtkkCard).filter(OtkkCard.content.is_(None)).delete()


def get_card_from_db(card_id: str, *, include_content: bool = False) -> dict[str, Any] | None:
    cid = str(card_id).strip()
    if not cid:
        return None
    init_db()
    _ensure_otkk_schema()
    with get_session() as session:
        row = session.get(OtkkCard, cid)
        if not row:
            return None
        return _row_to_dict(row, include_content=include_content)



def upsert_card_content(parsed: dict[str, Any], *, file_name: str | None = None) -> dict[str, Any]:
    """Полная перезапись структуры карты в БД (как в исходном .doc)."""
    cid = str(parsed.get("id") or "").strip()
    if not cid:
        raise ValueError("В распарсенной карте нет id")
    fname = Path(str(file_name or parsed.get("file") or "")).name.strip()
    if not fname:
        raise ValueError("Не задано имя файла карты")

    content = {
        "id": cid,
        "code": parsed.get("code") or "",
        "title": parsed.get("title") or "",
        "file": fname,
        "rows": sanitize_otkk_rows(parsed.get("rows") or []),
        "signature": parsed.get("signature"),
        "plain_text": parsed.get("plain_text") or "",
    }

    init_db()
    _ensure_otkk_schema()
    with get_session() as session:
        row = session.get(OtkkCard, cid)
        if row:
            row.file_name = fname
            row.code = content["code"]
            row.title = content["title"]
            row.content = content
        else:
            session.add(
                OtkkCard(
                    id=cid,
                    file_name=fname,
                    code=content["code"],
                    title=content["title"],
                    content=content,
                )
            )
    return {"id": cid, "file": fname, "code": content["code"], "title": content["title"]}


def _seed_otkk_card(
    loader: str,
    *,
    overwrite: bool = False,
) -> dict[str, Any]:
    if loader == "otkk-1":
        from sk_reporter.otkk1_data import otkk1_parsed

        parsed = otkk1_parsed()
    elif loader == "otkk-2":
        from sk_reporter.otkk2_data import otkk2_parsed

        parsed = otkk2_parsed()
    else:
        raise ValueError(f"Неизвестная эталонная карта: {loader}")

    cid = parsed["id"]
    if not overwrite:
        existing = get_card_from_db(cid, include_content=True)
        if existing and existing.get("has_content"):
            return {"id": cid, "skipped": True, "reason": "already in db"}
    result = upsert_card_content(parsed, file_name=parsed["file"])
    result["rows"] = len(parsed.get("rows") or [])
    result["seeded"] = True
    return result


def seed_otkk1(*, overwrite: bool = False) -> dict[str, Any]:
    """Залить эталон ОТКК-1 (6 пунктов карты) в PostgreSQL."""
    return _seed_otkk_card("otkk-1", overwrite=overwrite)


def seed_otkk2(*, overwrite: bool = False) -> dict[str, Any]:
    """Залить эталон ОТКК-2 (6 пунктов карты) в PostgreSQL."""
    return _seed_otkk_card("otkk-2", overwrite=overwrite)


def import_document_to_db(
    source: Path,
    *,
    copy_to_tk_dir: bool = False,
    tk_root: Path | None = None,
) -> dict[str, Any]:
    from sk_reporter.otkk_parser import parse_otkk_document

    source = Path(source)
    if not source.is_file():
        raise FileNotFoundError(source)
    parsed = parse_otkk_document(source)
    if not parsed.get("id"):
        raise ValueError(f"Не удалось определить id ОТКК из файла: {source.name}")

    file_name = source.name
    if copy_to_tk_dir:
        root = tk_root or tk_dir()
        root.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, root / file_name)

    result = upsert_card_content(parsed, file_name=file_name)
    result["rows"] = len(parsed.get("rows") or [])
    return result

