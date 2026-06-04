"""
Rule-based переключение руководителя в ежедневных отчётах (блок 2 UI).

Шапка — короткая должность + ФИО; подвал — полная должность + ФИО.
Обрабатываются все подходящие ячейки в верхней/нижней зоне таблицы
(сетки GRID_COLS_6 / GRID_COLS_7 с ghost-колонкой).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

from docx import Document
from docx.table import _Cell

from sk_reporter.template_layout import _main_table_indices

LeaderId = Literal["aniskov", "mandzhiev"]

LEADER_TARGETS: dict[LeaderId, dict[str, str]] = {
    "aniskov": {
        "header_title": "Руководитель",
        "header_fio": "Аниськов Владимир Иванович",
        "footer_role": "Руководитель проекта СК",
        "footer_fio": "Аниськов Владимир Иванович",
    },
    "mandzhiev": {
        "header_title": "И.О. Руководителя",
        "header_fio": "Манджиев Игорь Александрович",
        "footer_role": "И.О. Руководителя проекта СК",
        "footer_fio": "Манджиев Игорь Александрович",
    },
}

# Варианты должности в подвале (для замены в run'ах)
_FOOTER_ROLE_VARIANTS = [
    "И.О. Руководителя проекта СК",
    "И.о. Руководителя проекта СК",
    "И.о. руководителя проекта СК",
    "Руководитель проекта СК",
]

_HEADER_TITLE_VARIANTS = [
    "И.О. Руководителя",
    "И.о. Руководителя",
    "И.о. руководителя",
    "Руководитель",
]

_FIO_REGEXES = [
    re.compile(r"Аниськов\s+В(?:ладимир)?\s+И(?:ванович)?\.?", re.I),
    re.compile(r"Манджиев\s+И(?:горь)?\s+А(?:лександрович)?\.?", re.I),
    re.compile(r"Аниськов\s+В\.?\s*И\.?", re.I),
    re.compile(r"Манджиев\s+И\.?\s*А\.?", re.I),
]

HEADER_ZONE_ROWS = 14
FOOTER_ZONE_ROWS = 14


@dataclass(frozen=True)
class CellSlot:
    row: int
    col: int


@dataclass
class LeaderSlots:
    header_title: CellSlot
    header_fio: CellSlot
    footer_role: CellSlot
    footer_fio: CellSlot


def _norm(text: str) -> str:
    return " ".join(text.split()).strip()


def _cell_text(cell: _Cell) -> str:
    return _norm(cell.text)


def _is_leader_fio(text: str) -> bool:
    t = text.lower()
    if "аниськов" in t or "манджиев" in t:
        return True
    return any(r.search(text) for r in _FIO_REGEXES)


def _is_header_title(text: str) -> bool:
    t = text.lower()
    if "руководител" not in t:
        return False
    if "лицеванов" in t or "проекта" in t:
        return False
    return True


def _is_footer_role(text: str) -> bool:
    t = text.lower()
    if "руководител" not in t or "проекта" not in t:
        return False
    if "лицеванов" in t:
        return False
    return True


def _skip_cell(text: str) -> bool:
    t = text.lower()
    return not text or "лицеванов" in t


def _right_band_cells(row, min_col: int = 2) -> list[tuple[int, _Cell]]:
    """Уникальные непустые ячейки правой части строки (подписи)."""
    cells = list(enumerate(row.cells))
    if not cells:
        return []
    start = max(min_col, len(cells) - 4)
    out: list[tuple[int, _Cell]] = []
    seen: set[int] = set()
    for ci, cell in cells:
        if ci < start:
            continue
        tc_id = id(cell._tc)
        if tc_id in seen:
            continue
        seen.add(tc_id)
        if _cell_text(cell):
            out.append((ci, cell))
    return out


def find_leader_slots(doc: Document) -> LeaderSlots | None:
    """Диагностика: первая найденная четвёрка слотов (как раньше)."""
    indices = _main_table_indices(doc)
    if not indices:
        return None
    table = doc.tables[indices[0]]
    nrows = len(table.rows)
    if nrows < 12:
        return None

    header_title: CellSlot | None = None
    header_fio: CellSlot | None = None
    for ri in range(min(HEADER_ZONE_ROWS, nrows)):
        titles: list[int] = []
        fios: list[int] = []
        for ci, cell in _right_band_cells(table.rows[ri]):
            text = _cell_text(cell)
            if _is_leader_fio(text):
                fios.append(ci)
            elif _is_header_title(text):
                titles.append(ci)
        if titles and fios:
            header_title = CellSlot(ri, titles[-1])
            header_fio = CellSlot(ri, fios[-1])
            break

    if header_title is None or header_fio is None:
        return None

    footer_role: CellSlot | None = None
    footer_fio: CellSlot | None = None
    scan_from = max(nrows - FOOTER_ZONE_ROWS, header_title.row + 1)
    for ri in range(nrows - 1, scan_from - 1, -1):
        for ci, cell in _right_band_cells(table.rows[ri]):
            text = _cell_text(cell)
            if footer_fio is None and _is_leader_fio(text):
                footer_fio = CellSlot(ri, ci)
            if footer_role is None and _is_footer_role(text):
                footer_role = CellSlot(ri, ci)

    if footer_role is None or footer_fio is None:
        return None

    return LeaderSlots(
        header_title=header_title,
        header_fio=header_fio,
        footer_role=footer_role,
        footer_fio=footer_fio,
    )


def _write_cell_text(cell: _Cell, new_text: str) -> bool:
    if _cell_text(cell) == new_text:
        return False
    if cell.paragraphs:
        cell.paragraphs[0].text = new_text
        for para in cell.paragraphs[1:]:
            para.text = ""
    else:
        cell.text = new_text
    return True


def _replace_in_runs(cell: _Cell, old_variants: list[str], new_text: str) -> bool:
    changed = False
    for para in cell.paragraphs:
        for run in para.runs:
            original = run.text
            updated = original
            for old in old_variants:
                if old in updated:
                    updated = updated.replace(old, new_text)
                else:
                    m = re.search(re.escape(old), updated, re.I)
                    if m:
                        updated = updated[: m.start()] + new_text + updated[m.end() :]
            if updated != original:
                run.text = updated
                changed = True
    return changed


def _replace_title_in_cell(cell: _Cell, target: str) -> bool:
    if _cell_text(cell) == target:
        return False
    if _replace_in_runs(cell, _HEADER_TITLE_VARIANTS, target):
        return True
    if _is_header_title(_cell_text(cell)):
        return _write_cell_text(cell, target)
    return False


def _replace_footer_role_in_cell(cell: _Cell, target: str) -> bool:
    if _cell_text(cell) == target:
        return False
    if _replace_in_runs(cell, _FOOTER_ROLE_VARIANTS, target):
        return True
    if _is_footer_role(_cell_text(cell)):
        return _write_cell_text(cell, target)
    return False


def _replace_fio_in_cell(cell: _Cell, target_fio: str) -> bool:
    if _cell_text(cell) == target_fio:
        return False
    changed = False
    for para in cell.paragraphs:
        for run in para.runs:
            new = run.text
            for rx in _FIO_REGEXES:
                new = rx.sub(target_fio, new)
            if new != run.text:
                run.text = new
                changed = True
    if changed:
        return True
    if _is_leader_fio(_cell_text(cell)) or not _cell_text(cell):
        return _write_cell_text(cell, target_fio)
    return False


def _apply_header_zone(table, leader: LeaderId) -> int:
    targets = LEADER_TARGETS[leader]
    nrows = len(table.rows)
    changes = 0
    seen_tc: set[int] = set()

    for ri in range(min(HEADER_ZONE_ROWS, nrows)):
        for _ci, cell in _right_band_cells(table.rows[ri]):
            tc_id = id(cell._tc)
            if tc_id in seen_tc:
                continue
            text = _cell_text(cell)
            if _skip_cell(text):
                continue

            if _is_leader_fio(text):
                seen_tc.add(tc_id)
                if _replace_fio_in_cell(cell, targets["header_fio"]):
                    changes += 1
            elif _is_header_title(text):
                seen_tc.add(tc_id)
                if _replace_title_in_cell(cell, targets["header_title"]):
                    changes += 1

    return changes


def _apply_footer_zone(table, leader: LeaderId) -> int:
    targets = LEADER_TARGETS[leader]
    nrows = len(table.rows)
    scan_from = max(0, nrows - FOOTER_ZONE_ROWS)
    changes = 0
    seen_tc: set[int] = set()

    for ri in range(scan_from, nrows):
        for _ci, cell in _right_band_cells(table.rows[ri]):
            tc_id = id(cell._tc)
            if tc_id in seen_tc:
                continue
            text = _cell_text(cell)
            if _skip_cell(text):
                continue

            # Сначала полная должность (чтобы не спутать с ФИО в одной ячейке)
            if _is_footer_role(text):
                seen_tc.add(tc_id)
                if _replace_footer_role_in_cell(cell, targets["footer_role"]):
                    changes += 1
                continue

            if _is_leader_fio(text):
                seen_tc.add(tc_id)
                if _replace_fio_in_cell(cell, targets["footer_fio"]):
                    changes += 1

    return changes


def switch_leader_in_docx(filepath: str, leader: LeaderId) -> tuple[bool, str, int]:
    path = Path(filepath)
    try:
        doc = Document(str(path))
        indices = _main_table_indices(doc)
        if not indices:
            return False, "нет таблиц", 0
        table = doc.tables[indices[0]]

        if find_leader_slots(doc) is None:
            return False, "не найдены ячейки руководителя", 0

        changes = _apply_header_zone(table, leader) + _apply_footer_zone(table, leader)

        if changes == 0:
            return True, "уже нужный руководитель", 0

        doc.save(str(path))
        return True, f"замен: {changes}", changes
    except Exception as e:
        return False, str(e), 0


def switch_leader_batch(filepaths: list[str], leader: LeaderId) -> list[dict]:
    results = []
    for fp in filepaths:
        name = Path(fp).name
        ok, msg, n = switch_leader_in_docx(fp, leader)
        results.append({"filename": name, "ok": ok, "msg": msg, "changes": n})
    return results
