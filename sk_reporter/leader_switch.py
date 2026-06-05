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
    "Руководитель проекта",
]

_HEADER_TITLE_VARIANTS = [
    "И.О. Руководителя",
    "И.о. Руководителя",
    "И.о. руководителя",
    "Руководитель проекта СК",
    "Руководитель проекта",
    "Руководитель",
]

_FIO_REGEXES = [
    re.compile(r"Аниськов\s+В(?:ладимир)?\s+И(?:ванович)?\.?", re.I),
    re.compile(r"Манджиев\s+И(?:горь)?\s+А(?:лександрович)?\.?", re.I),
    re.compile(r"Аниськов\s+В\.?\s*И\.?", re.I),
    re.compile(r"Манджиев\s+И\.?\s*А\.?", re.I),
]

# Должность в подвале (в т.ч. разбитая по run/абзацам)
_FOOTER_ROLE_RE = re.compile(
    r"(?:и\.?\s*о\.?\s*)?руководител\w*\s+проект\w*(?:\s+с\.?\s*к\.?)?",
    re.I,
)

HEADER_ZONE_ROWS = 14
FOOTER_ZONE_ROWS = 18


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
    """Короткая должность в шапке (без «проекта»)."""
    t = text.lower()
    if "руководител" not in t:
        return False
    if "лицеванов" in t or "проекта" in t:
        return False
    return True


def _is_header_title_row_cell(text: str) -> bool:
    """Ячейка должности в строке шапки (в т.ч. ошибочно «…проекта СК» в GRID_COLS_7)."""
    t = text.lower()
    if "руководител" not in t:
        return False
    if "лицеванов" in t:
        return False
    if _is_leader_fio(text) and not _FOOTER_ROLE_RE.search(text):
        return False
    return True


def _is_footer_role(text: str) -> bool:
    t = text.lower()
    if "лицеванов" in t:
        return False
    if _FOOTER_ROLE_RE.search(t):
        return True
    return "руководител" in t and "проект" in t


def _is_footer_role_paragraph(text: str) -> bool:
    """Абзац с полной должностью (не чистое ФИО)."""
    t = _norm(text)
    if not t or "лицеванов" in t.lower():
        return False
    if _is_leader_fio(t) and not _FOOTER_ROLE_RE.search(t):
        return False
    return _is_footer_role(t)


def _is_fio_paragraph(text: str) -> bool:
    t = _norm(text)
    return bool(t) and _is_leader_fio(t) and not _is_footer_role_paragraph(t)


def _skip_cell(text: str) -> bool:
    t = text.lower()
    return not text or "лицеванов" in t


def _zone_cells(row) -> list[tuple[int, _Cell]]:
    """Все непустые уникальные ячейки строки (для подвала — без обрезки по col)."""
    out: list[tuple[int, _Cell]] = []
    seen: set[int] = set()
    for ci, cell in enumerate(row.cells):
        tc_id = id(cell._tc)
        if tc_id in seen:
            continue
        seen.add(tc_id)
        if _cell_text(cell):
            out.append((ci, cell))
    return out


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
            elif _is_header_title_row_cell(text):
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


def _replace_in_runs_paragraph(para, old_variants: list[str], new_text: str) -> bool:
    changed = False
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


def _replace_in_runs(cell: _Cell, old_variants: list[str], new_text: str) -> bool:
    changed = False
    for para in cell.paragraphs:
        if _replace_in_runs_paragraph(para, old_variants, new_text):
            changed = True
    return changed


def _replace_footer_role_in_paragraph(para, target: str) -> bool:
    ptext = _norm(para.text)
    if not ptext or ptext == target or "лицеванов" in ptext.lower():
        return False
    if not _is_footer_role_paragraph(ptext):
        return False
    if _replace_in_runs_paragraph(para, _FOOTER_ROLE_VARIANTS, target):
        return True
    new_pt = _FOOTER_ROLE_RE.sub(target, para.text, count=1)
    if new_pt != para.text:
        para.text = new_pt
        return True
    if ptext != target:
        para.text = target
        return True
    return False


def _replace_title_in_cell(cell: _Cell, target: str) -> bool:
    if _cell_text(cell) == target:
        return False
    if _replace_in_runs(cell, _HEADER_TITLE_VARIANTS, target):
        return True
    if _is_header_title_row_cell(_cell_text(cell)):
        return _write_cell_text(cell, target)
    return False


def _replace_footer_role_in_cell(cell: _Cell, target: str) -> bool:
    if _cell_text(cell) == target:
        return False
    changed = False
    for para in cell.paragraphs:
        if _replace_footer_role_in_paragraph(para, target):
            changed = True
    if changed:
        return True
    if _is_footer_role(_cell_text(cell)):
        return _write_cell_text(cell, target)
    return False


def _replace_fio_in_paragraph(para, target_fio: str) -> bool:
    if _norm(para.text) == target_fio:
        return False
    changed = False
    for run in para.runs:
        new = run.text
        for rx in _FIO_REGEXES:
            new = rx.sub(target_fio, new)
        if new != run.text:
            run.text = new
            changed = True
    if changed:
        return True
    if _is_leader_fio(_norm(para.text)):
        para.text = target_fio
        return True
    return False


def _replace_fio_in_cell(cell: _Cell, target_fio: str) -> bool:
    if _cell_text(cell) == target_fio:
        return False
    changed = False
    for para in cell.paragraphs:
        if _replace_fio_in_paragraph(para, target_fio):
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
            elif _is_header_title_row_cell(text):
                seen_tc.add(tc_id)
                if _replace_title_in_cell(cell, targets["header_title"]):
                    changes += 1

    return changes


def _row_footer_role_cells(row) -> list[_Cell]:
    """Ячейки строки, в сумме дающие должность в подвале (в т.ч. разбитую по col)."""
    candidates: list[tuple[int, _Cell, str]] = []
    for ci, cell in _zone_cells(row):
        text = _cell_text(cell)
        if not text or "лицеванов" in text.lower():
            continue
        tl = text.lower()
        if _is_leader_fio(text) and not _FOOTER_ROLE_RE.search(text):
            continue
        if "руководител" in tl or ("проект" in tl and "ск" in tl.replace(".", "")):
            candidates.append((ci, cell, text))
    if not candidates:
        return []
    combined = _norm(" ".join(t for _, _, t in candidates))
    if not (_is_footer_role(combined) or _FOOTER_ROLE_RE.search(combined)):
        return []
    return [cell for _, cell, _ in candidates]


def _apply_footer_role_row(row, target: str) -> bool:
    cells = _row_footer_role_cells(row)
    if not cells:
        return False
    if len(cells) == 1:
        return _replace_footer_role_in_cell(cells[0], target)
    changed = False
    if _write_cell_text(cells[0], target):
        changed = True
    for cell in cells[1:]:
        if _cell_text(cell) and not _is_leader_fio(_cell_text(cell)):
            if _write_cell_text(cell, ""):
                changed = True
    return changed


def _apply_footer_zone(table, leader: LeaderId) -> int:
    targets = LEADER_TARGETS[leader]
    nrows = len(table.rows)
    scan_from = max(0, nrows - FOOTER_ZONE_ROWS)
    changes = 0

    for ri in range(scan_from, nrows):
        row = table.rows[ri]
        row_changed = False

        for _ci, cell in _zone_cells(row):
            for para in cell.paragraphs:
                ptext = _norm(para.text)
                if not ptext or "лицеванов" in ptext.lower():
                    continue
                if _is_fio_paragraph(ptext):
                    if _replace_fio_in_paragraph(para, targets["footer_fio"]):
                        row_changed = True

        if _apply_footer_role_row(row, targets["footer_role"]):
            row_changed = True
        else:
            for _ci, cell in _zone_cells(row):
                text = _cell_text(cell)
                if _skip_cell(text):
                    continue
                if _is_footer_role(text):
                    if _replace_footer_role_in_cell(cell, targets["footer_role"]):
                        row_changed = True
                elif _is_leader_fio(text):
                    if _replace_fio_in_cell(cell, targets["footer_fio"]):
                        row_changed = True
                else:
                    for para in cell.paragraphs:
                        ptext = _norm(para.text)
                        if _is_footer_role_paragraph(ptext):
                            if _replace_footer_role_in_paragraph(
                                para, targets["footer_role"]
                            ):
                                row_changed = True

        if row_changed:
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

        slots = find_leader_slots(doc)
        changes = _apply_header_zone(table, leader) + _apply_footer_zone(table, leader)

        if changes == 0:
            if slots is None:
                return False, "не найдены ячейки руководителя", 0
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
