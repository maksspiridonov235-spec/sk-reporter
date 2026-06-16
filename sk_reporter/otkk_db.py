"""Каталог ОТКК (технологические карты) в PostgreSQL."""

from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any

import yaml
from sqlalchemy import text

from sk_reporter.db.config import database_enabled, database_url
from sk_reporter.db.models import OtkkCard
from sk_reporter.db.session import get_session, init_db
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
            count = session.query(OtkkCard).count()
            with_content = session.query(OtkkCard).filter(OtkkCard.content.isnot(None)).count()
        return {
            "enabled": True,
            "configured": True,
            "count": count,
            "with_content": with_content,
            "ok": True,
        }
    except Exception as exc:
        return {"enabled": True, "configured": True, "count": 0, "with_content": 0, "ok": False, "error": str(exc)}


def _row_to_dict(row: OtkkCard, *, include_content: bool = False) -> dict[str, Any]:
    out: dict[str, Any] = {
        "id": row.id,
        "file": row.file_name,
        "code": row.code or "",
        "title": row.title or "",
        "has_content": row.content is not None,
    }
    if include_content and row.content is not None:
        out["content"] = row.content
    return out


def load_cards_from_db() -> list[dict[str, Any]]:
    init_db()
    _ensure_otkk_schema()
    with get_session() as session:
        rows = session.query(OtkkCard).order_by(OtkkCard.id).all()
        return [_row_to_dict(r) for r in rows]


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


def upsert_cards(cards: list[dict[str, Any]]) -> dict[str, Any]:
    if not cards:
        return {"upserted": 0}
    init_db()
    _ensure_otkk_schema()
    upserted = 0
    with get_session() as session:
        for card in cards:
            cid = str(card.get("id") or "").strip()
            fname = Path(str(card.get("file") or "")).name.strip()
            if not cid or not fname:
                continue
            row = session.get(OtkkCard, cid)
            if row:
                row.file_name = fname
            else:
                session.add(OtkkCard(id=cid, file_name=fname))
            upserted += 1
    return {"upserted": upserted, "total": len(cards)}


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
        "rows": parsed.get("rows") or [],
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


def import_document_to_db(
    source: Path,
    *,
    copy_to_tk_dir: bool = True,
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


def parse_manifest_cards(path: Path) -> list[dict[str, Any]]:
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    cards: list[dict[str, Any]] = []
    for card in data.get("cards") or []:
        cid = str(card.get("id") or "").strip()
        fname = Path(str(card.get("file") or "")).name.strip()
        if cid and fname:
            cards.append({"id": cid, "file": fname})
    return cards


def import_manifest_to_db(path: Path) -> dict[str, Any]:
    cards = parse_manifest_cards(path)
    if not cards:
        raise ValueError("В manifest.yaml нет записей cards")
    result = upsert_cards(cards)
    result["source"] = "manifest"
    return result


def scan_disk_and_upsert() -> dict[str, Any]:
    from sk_reporter.engineer.tk_catalog import list_tk_files

    cards = list_tk_files()
    if not cards:
        raise ValueError("В data/tk/ нет файлов .doc/.docx")
    result = upsert_cards(cards)
    result["source"] = "disk"
    return result


def import_all_documents_from_disk() -> dict[str, Any]:
    """Парсит .doc/.docx на диске в БД, если structured content ещё нет."""
    from sk_reporter.otkk_parser import otkk_id_from_path

    folder = tk_dir()
    if not folder.is_dir():
        return {"imported": [], "skipped": 0, "errors": []}

    init_db()
    _ensure_otkk_schema()
    imported: list[str] = []
    errors: list[dict[str, str]] = []
    skipped = 0

    with get_session() as session:
        rows = {r.id: r for r in session.query(OtkkCard).all()}

    for path in sorted(folder.iterdir()):
        if path.suffix.lower() not in {".doc", ".docx"}:
            continue
        card_id = otkk_id_from_path(path)
        if not card_id:
            continue
        row = rows.get(card_id)
        if row and row.content is not None:
            skipped += 1
            continue
        try:
            import_document_to_db(path, copy_to_tk_dir=False)
            imported.append(card_id)
        except Exception as exc:
            errors.append({"id": card_id, "file": path.name, "error": str(exc)})

    return {"imported": imported, "skipped": skipped, "errors": errors}


def bootstrap_otkk_on_startup() -> dict[str, Any] | None:
    """При пустой БД — каталог из manifest; затем content из .doc на диске сервера."""
    if not database_enabled():
        return None
    init_db()
    _ensure_otkk_schema()
    st = db_status()
    if not st.get("ok"):
        return {"ok": False, "error": st.get("error")}

    folder = tk_dir()
    catalog: dict[str, Any] | None = None
    if st.get("count", 0) == 0:
        manifest = folder / "manifest.yaml"
        if manifest.is_file():
            catalog = import_manifest_to_db(manifest)
        elif any(folder.glob("*.doc")) or any(folder.glob("*.DOC")) or any(folder.glob("*.docx")):
            catalog = scan_disk_and_upsert()

    docs = import_all_documents_from_disk()
    return {"catalog": catalog, "documents": docs}
