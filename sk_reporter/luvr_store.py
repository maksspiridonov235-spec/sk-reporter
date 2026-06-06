"""ЛУВР — лист учёта времени (xlsx → luvr.yaml)."""

from __future__ import annotations

import re
from datetime import date, datetime
from functools import lru_cache
from pathlib import Path
from typing import Any

import yaml
from openpyxl import load_workbook

from sk_reporter.paths import luvr_dir

_MONTH_SHEETS = ("Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь")
_MONTH_NUM = {name: i + 1 for i, name in enumerate(_MONTH_SHEETS)}


def _luvr_xlsx() -> Path:
    folder = luvr_dir()
    for name in ("ЛУВР.xlsx", "luvr.xlsx"):
        p = folder / name
        if p.is_file():
            return p
    matches = sorted(folder.glob("*.xlsx"))
    if matches:
        return matches[0]
    raise FileNotFoundError(f"Нет xlsx в {folder}")


def _cell_date(v: Any) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    return None


def _norm_fio(v: Any) -> str:
    return " ".join(str(v or "").split())


def _count_days(values: list[Any]) -> dict[str, int]:
    out = {"present": 0, "half": 0, "other": 0}
    for v in values:
        if v is None or v == "":
            continue
        s = str(v).strip()
        if s == "1":
            out["present"] += 1
        elif s in {"0.5", "0,5"}:
            out["half"] += 1
        else:
            out["other"] += 1
    return out


def _norm_mark(v: Any) -> str:
    if v is None or v == "":
        return ""
    s = str(v).strip().replace(",", ".")
    if s == "1":
        return "1"
    if s == "0.5":
        return "0.5"
    return s


def _parse_sheet(ws) -> dict[str, Any] | None:
    rows = list(ws.iter_rows(values_only=True))
    hdr_idx = None
    for i, row in enumerate(rows):
        if row and row[1] == "ФИО":
            hdr_idx = i
            break
    if hdr_idx is None:
        return None

    title = ""
    for row in rows[: hdr_idx + 1]:
        if row and row[0] and isinstance(row[0], str) and "Лист учета" in row[0]:
            title = row[0].replace("\n", " ").strip()
            break

    day_row = rows[hdr_idx + 1] if hdr_idx + 1 < len(rows) else ()
    day_cols: list[dict[str, Any]] = []
    for col_idx, val in enumerate(day_row):
        d = _cell_date(val)
        if d:
            day_cols.append({"col": col_idx, "date": d.isoformat(), "day": d.day})

    people: list[dict[str, Any]] = []
    for row in rows[hdr_idx + 2 :]:
        if not row or not _norm_fio(row[1]):
            continue
        day_values = [row[c["col"]] if c["col"] < len(row) else None for c in day_cols]
        counts = _count_days(day_values)
        marks = [_norm_mark(v) for v in day_values]
        people.append(
            {
                "num": row[0],
                "fio": _norm_fio(row[1]),
                "position": str(row[2] or "").strip(),
                "nrs": str(row[3] or "").strip(),
                "specialty": str(row[4] or "").strip(),
                "days_present": counts["present"],
                "days_half": counts["half"],
                "days_marked": counts["present"] + counts["half"] + counts["other"],
                "marks": marks,
            }
        )

    month_num = _MONTH_NUM.get(ws.title)
    year = day_cols[0]["date"][:4] if day_cols else None
    return {
        "sheet": ws.title,
        "year": int(year) if year else None,
        "month": month_num,
        "title": title,
        "people_count": len(people),
        "days_in_sheet": len(day_cols),
        "days": [{"date": c["date"], "day": c["day"]} for c in day_cols],
        "people": people,
    }


def export_luvr() -> Path:
    xlsx = _luvr_xlsx()
    wb = load_workbook(xlsx, read_only=True, data_only=True)
    months: list[dict[str, Any]] = []
    contract = ""
    for sheet_name in wb.sheetnames:
        if sheet_name not in _MONTH_SHEETS:
            continue
        parsed = _parse_sheet(wb[sheet_name])
        if parsed and parsed["people_count"]:
            if not contract and parsed.get("title"):
                m = re.search(r"договор[^\d]*([\w.\-]+)", parsed["title"], re.I)
                if m:
                    contract = m.group(1)
            months.append(parsed)
    wb.close()

    payload = {
        "source": xlsx.name,
        "source_kb": round(xlsx.stat().st_size / 1024, 1),
        "contract": contract,
        "months": months,
    }
    out = luvr_dir() / "luvr.yaml"
    out.write_text(yaml.safe_dump(payload, allow_unicode=True, sort_keys=False), encoding="utf-8")
    load_luvr.cache_clear()
    return out


def _recalc_person_stats(person: dict[str, Any]) -> None:
    marks = person.get("marks") or []
    counts = _count_days(marks)
    person["days_present"] = counts["present"]
    person["days_half"] = counts["half"]
    person["days_marked"] = counts["present"] + counts["half"] + counts["other"]


def save_luvr(data: dict[str, Any] | None = None) -> Path:
    payload = data if data is not None else load_luvr()
    out = luvr_dir() / "luvr.yaml"
    out.write_text(yaml.safe_dump(payload, allow_unicode=True, sort_keys=False), encoding="utf-8")
    load_luvr.cache_clear()
    return out


def update_luvr_mark(sheet: str, person_idx: int, day_idx: int, mark: Any) -> dict[str, Any]:
    if sheet not in _MONTH_SHEETS:
        raise KeyError(f"Неизвестный лист: {sheet}")
    data = load_luvr()
    if not data.get("months"):
        raise FileNotFoundError("luvr.yaml пуст — сначала build_engineer_data.py --luvr")

    month = next((m for m in data["months"] if m.get("sheet") == sheet), None)
    if month is None:
        raise KeyError(f"Лист «{sheet}» не найден в luvr.yaml")

    people = month.get("people") or []
    if person_idx < 0 or person_idx >= len(people):
        raise IndexError("person_idx вне диапазона")

    days = month.get("days") or []
    if day_idx < 0 or day_idx >= len(days):
        raise IndexError("day_idx вне диапазона")

    person = people[person_idx]
    marks = list(person.get("marks") or [])
    while len(marks) < len(days):
        marks.append("")

    norm = _norm_mark(mark)
    marks[day_idx] = norm
    person["marks"] = marks
    _recalc_person_stats(person)
    save_luvr(data)

    return {
        "sheet": sheet,
        "person_idx": person_idx,
        "day_idx": day_idx,
        "mark": norm,
        "days_present": person["days_present"],
        "days_marked": person["days_marked"],
    }


@lru_cache(maxsize=1)
def load_luvr() -> dict[str, Any]:
    cache = luvr_dir() / "luvr.yaml"
    if cache.is_file():
        return yaml.safe_load(cache.read_text(encoding="utf-8")) or {}
    return {}


def _cache_has_grid(months: list[dict[str, Any]]) -> bool:
    if not months:
        return False
    m = months[0]
    if not m.get("days"):
        return False
    people = m.get("people") or []
    return bool(people and "marks" in people[0])


def luvr_planning_payload() -> dict[str, Any]:
    from sk_reporter.paths import repo_root

    folder = luvr_dir()
    xlsx_path = None
    try:
        xlsx_path = _luvr_xlsx()
    except FileNotFoundError:
        pass

    data = load_luvr()
    months = data.get("months") or []
    if xlsx_path is not None and (not months or not _cache_has_grid(months)):
        try:
            export_luvr()
            data = load_luvr()
            months = data.get("months") or []
        except Exception:
            pass

    yaml_path = luvr_dir() / "luvr.yaml"
    default_month = months[-1]["sheet"] if months else None

    return {
        "folder": str(folder.relative_to(repo_root())),
        "source": data.get("source") or (xlsx_path.name if xlsx_path else None),
        "source_kb": data.get("source_kb") or (round(xlsx_path.stat().st_size / 1024, 1) if xlsx_path else None),
        "contract": data.get("contract"),
        "months": months,
        "default_month": default_month,
        "cache_ready": bool(months),
        "cache_from_yaml": bool(months) and yaml_path.is_file(),
        "grid_ready": _cache_has_grid(months),
        "editable": _cache_has_grid(months),
        "xlsx_present": xlsx_path is not None,
    }


def luvr_month_payload(sheet: str) -> dict[str, Any]:
    data = load_luvr()
    for m in data.get("months") or []:
        if m.get("sheet") == sheet:
            return m
    raise KeyError(sheet)
