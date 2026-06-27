"""Импорт «Ядро.xlsx» в PostgreSQL: подрядчики и карточки проектов (шифр + объект)."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from sk_reporter.companies import COMPANIES
from sk_reporter.contractor_db import (
    _CONTRACTOR_ID_OVERRIDES,
    _slugify_id,
    ensure_contractor_schema,
)
from sk_reporter.core_xlsx import parse_core_contractors, parse_core_objects
from sk_reporter.db.config import database_enabled
from sk_reporter.db.models import Contractor, Project
from sk_reporter.db.session import get_session, init_db


def _require_database() -> None:
    if not database_enabled():
        raise RuntimeError("DATABASE_URL не задан")


def _template_stem_for_label(file_label: str) -> str:
    label = file_label.strip()
    for company in COMPANIES:
        if company.name.lower() == label.lower():
            return company.template_stem
    return label


def _contractor_id_for_row(row: dict[str, Any], *, used: set[str]) -> str:
    file_label = str(row.get("file_label") or "").strip()
    sub = str(row.get("sub_contractor") or "").strip()
    contact = str(row.get("contact_person") or "").strip()
    base = _CONTRACTOR_ID_OVERRIDES.get(file_label) or _slugify_id(file_label or row.get("gen_contractor") or "contractor")

    parts = [base]
    if sub and sub not in {"—", "-", "–"}:
        parts.append(_slugify_id(sub))

    cid = "-".join(p for p in parts if p)
    if cid not in used:
        used.add(cid)
        return cid

    if contact:
        token = re.sub(r"[^\w]+", "-", contact.lower()).strip("-")[:20]
        if token:
            cid = f"{cid}-{token}"
    n = 2
    while cid in used:
        cid = f"{base}-{n}"
        n += 1
    used.add(cid)
    return cid


def _contractor_payload(row: dict[str, Any], contractor_id: str) -> dict[str, Any]:
    file_label = str(row.get("file_label") or "").strip()
    gen = str(row.get("gen_contractor") or file_label).strip()
    return {
        "id": contractor_id,
        "name": gen or file_label,
        "template_stem": _template_stem_for_label(file_label),
        "file_label": file_label,
        "inspection_type": str(row.get("inspection_type") or "").strip(),
        "gen_contractor": gen,
        "sub_contractor": str(row.get("sub_contractor") or "").strip(),
        "contract_no": str(row.get("contract_no") or "").strip(),
        "contact_person": str(row.get("contact_person") or "").strip(),
        "contact_phone": str(row.get("contact_phone") or "").strip(),
        "contact_fax": str(row.get("contact_fax") or "").strip(),
        "contact_email": str(row.get("contact_email") or "").strip(),
        "extra_info": str(row.get("extra_info") or "").strip(),
        "note_discrepancy": str(row.get("note_discrepancy") or "").strip(),
        "is_active": True,
    }


def _project_stub_id(cipher: str, org_label: str, *, cipher_counts: dict[str, int]) -> str:
    cipher = cipher.strip()
    if cipher_counts.get(cipher, 0) <= 1:
        return cipher[:64]
    pid = f"{cipher}__{_slugify_id(org_label)}"[:64]
    return pid


def import_core_xlsx_to_db(path: Path | str) -> dict[str, Any]:
    _require_database()
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(path)

    contractor_rows = parse_core_contractors(path)
    object_rows = parse_core_objects(path)

    init_db()
    ensure_contractor_schema()

    used_ids: set[str] = set()
    contractor_payloads: list[dict[str, Any]] = []
    for row in contractor_rows:
        cid = _contractor_id_for_row(row, used=used_ids)
        contractor_payloads.append(_contractor_payload(row, cid))

    cipher_counts: dict[str, int] = {}
    for obj in object_rows:
        c = str(obj["cipher"]).strip()
        cipher_counts[c] = cipher_counts.get(c, 0) + 1

    contractors_upserted = 0
    projects_upserted = 0
    projects_skipped_content = 0

    with get_session() as session:
        by_label: dict[str, list[Contractor]] = {}
        for payload in contractor_payloads:
            row = session.get(Contractor, payload["id"])
            if row:
                for key, val in payload.items():
                    if key != "id":
                        setattr(row, key, val)
            else:
                session.add(Contractor(**payload))
            contractors_upserted += 1

        session.flush()

        for payload in contractor_payloads:
            row = session.get(Contractor, payload["id"])
            if row:
                by_label.setdefault(row.file_label or payload["file_label"], []).append(row)

        def resolve_contractor_id(org_label: str) -> str | None:
            candidates = by_label.get(org_label) or []
            if not candidates:
                for payload in contractor_payloads:
                    if payload["file_label"] == org_label:
                        return payload["id"]
                return _CONTRACTOR_ID_OVERRIDES.get(org_label) or _slugify_id(org_label)
            for c in candidates:
                sub = (c.sub_contractor or "").strip()
                if not sub or sub in {"—", "-"}:
                    return c.id
            return candidates[0].id

        for obj in object_rows:
            org = str(obj["org_label"]).strip()
            cipher = str(obj["cipher"]).strip()
            object_name = str(obj["object_name"]).strip() or cipher
            contractor_id = resolve_contractor_id(org)
            if not contractor_id:
                raise ValueError(f"Не найден подрядчик для организации {org!r}")

            pid = _project_stub_id(cipher, org, cipher_counts=cipher_counts)
            existing = session.get(Project, pid)
            if existing and existing.content is not None:
                projects_skipped_content += 1
                if not existing.contractor_id:
                    existing.contractor_id = contractor_id
                continue

            fields = {
                "title": cipher,
                "object_name": object_name,
                "contractor_id": contractor_id,
                "vor_file": "",
                "tl_file": "",
                "content": None,
                "is_active": True,
            }
            if existing:
                for key, val in fields.items():
                    setattr(existing, key, val)
            else:
                session.add(Project(id=pid, **fields))
            projects_upserted += 1

    return {
        "source": str(path),
        "contractors_upserted": contractors_upserted,
        "projects_upserted": projects_upserted,
        "projects_skipped_with_content": projects_skipped_content,
        "contractors_total": len(contractor_payloads),
        "projects_total": len(object_rows),
    }
