"""Подрядчики в PostgreSQL (источник имён — болванки data/templates/)."""

from __future__ import annotations

import re
from typing import Any

from sqlalchemy import func

from sk_reporter.companies import COMPANIES, Company
from sk_reporter.db.config import database_enabled
from sk_reporter.db.models import Contractor, Project
from sk_reporter.db.session import get_session, init_db
from sk_reporter.paths import templates_dir

# Пилот; остальных подтянем кнопкой «из болванок»
_CONTRACTOR_ID_OVERRIDES: dict[str, str] = {
    "Евракор": "evrakor",
    "ОЗОТОБОС": "ozotobos",
    "Геодезический контроль": "geodez-kontrol",
    "Лесные технологии": "lesnye-tehnologii",
    "РНГМ-ГРУПП": "rngm-grupp",
    "Стройфинансгрупп": "stroyfinansgrupp",
}


def _slugify_id(text: str) -> str:
    raw = text.strip().lower()
    raw = raw.replace("«", "").replace("»", "").replace("\"", "")
    raw = re.sub(r"[^\w\-а-яё]+", "-", raw, flags=re.I)
    raw = re.sub(r"-+", "-", raw).strip("-")
    return raw or "contractor"


def contractor_id_for_company(company: Company) -> str:
    return _CONTRACTOR_ID_OVERRIDES.get(company.name) or _slugify_id(company.template_stem)


def _row_to_dict(row: Contractor, *, projects_count: int = 0) -> dict[str, Any]:
    return {
        "id": row.id,
        "name": row.name,
        "template_stem": row.template_stem or "",
        "is_active": bool(row.is_active),
        "projects_count": projects_count,
        "template_exists": (templates_dir() / f"{row.template_stem}.docx").is_file()
        if row.template_stem
        else False,
    }


def db_status() -> dict[str, Any]:
    if not database_enabled():
        return {
            "enabled": False,
            "configured": False,
            "count": 0,
            "ok": False,
            "error": "DATABASE_URL не задан",
        }
    try:
        init_db()
        with get_session() as session:
            count = session.query(Contractor).filter(Contractor.is_active.is_(True)).count()
        return {"enabled": True, "configured": True, "count": count, "ok": True}
    except Exception as exc:
        return {"enabled": True, "configured": True, "count": 0, "ok": False, "error": str(exc)}


def list_contractors(*, active_only: bool = True) -> list[dict[str, Any]]:
    init_db()
    with get_session() as session:
        q = session.query(Contractor)
        if active_only:
            q = q.filter(Contractor.is_active.is_(True))
        rows = q.order_by(Contractor.name).all()
        counts: dict[str, int] = dict(
            session.query(Project.contractor_id, func.count(Project.id))
            .filter(Project.is_active.is_(True))
            .group_by(Project.contractor_id)
            .all()
        )
        return [_row_to_dict(r, projects_count=counts.get(r.id, 0)) for r in rows]


def upsert_contractor(
    contractor_id: str,
    *,
    name: str,
    template_stem: str,
    is_active: bool = True,
) -> dict[str, Any]:
    init_db()
    with get_session() as session:
        row = session.get(Contractor, contractor_id)
        payload = {
            "name": name.strip(),
            "template_stem": template_stem.strip() or name.strip(),
            "is_active": is_active,
        }
        if not payload["name"]:
            raise ValueError("Название подрядчика обязательно")
        if row:
            for key, val in payload.items():
                setattr(row, key, val)
        else:
            row = Contractor(id=contractor_id, **payload)
            session.add(row)
        session.flush()
        return _row_to_dict(row)


def seed_contractor_company(company: Company, *, contractor_id: str | None = None) -> dict[str, Any]:
    cid = contractor_id or contractor_id_for_company(company)
    stem = company.template_stem
    tpl = templates_dir() / f"{stem}.docx"
    if not tpl.is_file():
        return {"id": cid, "seeded": False, "reason": f"нет болванки {stem}.docx"}
    row = upsert_contractor(cid, name=company.name, template_stem=stem)
    return {"id": cid, "seeded": True, "name": row["name"]}


def seed_evrakor() -> dict[str, Any]:
    company = next((c for c in COMPANIES if c.name == "Евракор"), None)
    if not company:
        return {"seeded": False, "reason": "Евракор не найден в companies.py"}
    return seed_contractor_company(company, contractor_id="evrakor")


def seed_contractors_from_templates(*, only_with_template: bool = True) -> dict[str, Any]:
    """Завести подрядчиков по всем COMPANIES, для которых есть .docx в data/templates/."""
    seeded: list[str] = []
    skipped: list[dict[str, str]] = []
    for company in COMPANIES:
        stem = company.template_stem
        if only_with_template and not (templates_dir() / f"{stem}.docx").is_file():
            skipped.append({"name": company.name, "reason": "нет болванки"})
            continue
        result = seed_contractor_company(company)
        if result.get("seeded"):
            seeded.append(result["id"])
        else:
            skipped.append({"name": company.name, "reason": result.get("reason", "пропуск")})
    return {"seeded": seeded, "seeded_count": len(seeded), "skipped": skipped}
