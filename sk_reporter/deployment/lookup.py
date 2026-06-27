"""Справочники для генерации расстановки: personnel (PostgreSQL), подрядчики, описания должностей."""

from __future__ import annotations

import json
import re
from functools import lru_cache
from typing import Any

from sk_reporter.contractor_db import list_contractors
from sk_reporter.paths import data_dir
from sk_reporter.personnel_store import load_people, _normalize_fio

DEFAULT_DESC = (
    "Контроль по сварочному производству\n"
    "Строительный контроль по общестроительным работам"
)
DEFAULT_REZHIM = "Инспекционный контроль, Проверка ИТД"

_FALLBACK_CONTRACTOR_MAP: list[tuple[list[str], str]] = [
    (["НГСК", "НОВАЯ ГАЗОВАЯ"], "ООО «Новая Газовая Строительная Компания»"),
    (["ЛЕСНЫЕ"], "ООО «Лесные Технологии»"),
    (["ЮГРАНЕФТЕ", "ЮГРАНЕФТЕСТРОЙ"], "ООО «ЮграНефтеСтрой»"),
    (["НЕФТЕСПЕЦСТРОЙ", "НСС"], "ООО «НефтеСпецСтрой»"),
    (["ЕВРАКОР"], "АО «Евракор»"),
    (["ТРУБОПРОВОДСЕРВИС"], "ООО ЭПЦ «Трубопроводсервис»"),
    (["ТЭКПРО"], "ООО «ТЭКПРО»"),
    (["СИБИТЕК"], "АО «Сибитек»"),
    (["ЭНЕРГОСТРОЙМОНТАЖ"], "ООО «ЭнергоСтройМонтаж»"),
    (["СТРОЙФИНАНСГРУПП"], "ООО «СтройФинансГрупп»"),
    (["НИПИ", "НЕФТЕГАЗПРОЕКТ"], "ООО НИПИ «Нефтегазпроект»"),
    (["РНГМ"], "АО «РНГМ-ГРУПП»"),
    (["ТЮМЕНЬВТОРСЫРЬЕ", "ТВС"], "ООО «ТюменьВторСырье»"),
    (["ТЮМЕНЬГЕОКОМ", "ТГК"], "ООО «ТюменьГеоКом»"),
    (["УРАЛГЕОГРУПП", "УГГ"], "ООО «УралГеоГрупп»"),
    (["ЮГРАГИДРОСТРОЙ", "ЮГС"], "ООО «Юграгидрострой»"),
    (["ЮГОРСКИЙ ПРОЕКТНЫЙ", "ЮПИ"], "ООО «Югорский проектный институт»"),
    (["РОСЭКСПО"], "ООО «РОСЭКСПО»"),
]


def _keywords_from_text(text: str) -> list[str]:
    upper = text.upper()
    parts = re.findall(r"[A-ZА-ЯЁ0-9]{3,}", upper)
    return parts


@lru_cache(maxsize=1)
def _contractor_rules() -> list[tuple[list[str], str]]:
    rules: list[tuple[list[str], str]] = []
    try:
        for c in list_contractors(active_only=True):
            canonical = (c.get("gen_contractor") or c.get("name") or "").strip()
            if not canonical:
                continue
            keywords: set[str] = set()
            if c.get("file_label"):
                keywords.add(c["file_label"].upper())
            for kw in _keywords_from_text(canonical):
                keywords.add(kw)
            if c.get("name"):
                for kw in _keywords_from_text(c["name"]):
                    keywords.add(kw)
            if keywords:
                rules.append((sorted(keywords, key=len, reverse=True), canonical))
    except Exception:
        pass
    rules.extend(_FALLBACK_CONTRACTOR_MAP)
    return rules


def normalize_contractor(name: str) -> str:
    if not name:
        return name
    upper = name.upper()
    for keywords, canonical in _contractor_rules():
        if any(kw in upper for kw in keywords):
            return canonical
    return name


@lru_cache(maxsize=1)
def _position_descriptions() -> dict[str, str]:
    path = data_dir() / "planning" / "position_descriptions.json"
    if not path.is_file():
        return {}
    try:
        rows = json.loads(path.read_text(encoding="utf-8"))
        return {
            str(r["dolzhnost"]).strip(): str(r.get("opisanie") or "").strip()
            for r in rows
            if r.get("dolzhnost")
        }
    except Exception:
        return {}


@lru_cache(maxsize=1)
def personnel_by_fio() -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for p in load_people():
        fio = _normalize_fio(p.get("fio") or "")
        if fio:
            result[fio] = p
    return result


def known_fio_set() -> set[str]:
    return set(personnel_by_fio().keys())


def load_desc_map() -> dict[str, str]:
    """{ФИО: описание действий} через personnel.position → position_descriptions."""
    pos_desc = _position_descriptions()
    result: dict[str, str] = {}
    for fio, person in personnel_by_fio().items():
        position = (person.get("position") or "").strip()
        result[fio] = pos_desc.get(position) or DEFAULT_DESC
    return result


def load_sprav_dict() -> dict[str, dict[str, str]]:
    """Справочник для расстановки: {fio: {dolzhnost, telefon, rezhim}}."""
    result: dict[str, dict[str, str]] = {}
    for fio, person in personnel_by_fio().items():
        result[fio] = {
            "dolzhnost": (person.get("position") or "").strip(),
            "telefon": (person.get("phone") or "").strip(),
            "rezhim": (person.get("control_mode") or "").strip() or DEFAULT_REZHIM,
        }
    return result
