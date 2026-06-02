"""
Применяет фиксированную сетку столбцов и высоту строк из шаблона ко всем таблицам документа.
Не зависит от количества строк или ячеек в документе.

Использование: python3 apply_template_layout.py <document.docx>
"""

import sys
import os
from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

TEMPLATE_PATH = (
    Path(__file__).parent
    / "contractor_report"
    / "болванки (шаблоны не вырезать только копировать)"
    / "Ежедневный отчет Шаблон.docx"
)

# ─── Геометрия сетки ──────────────────────────────────────────────────────────
#
# ПРАВИЛО: оба GRID_COLS должны содержать ровно 6 чисел с суммой 10830.
# tblW всегда = 10830 — выставляется жёстко, не вычисляется из ширины страницы.
#
# История проблемы:
#   Код генерации отчётов ранее создавал 7 gridCol вместо 6, вставляя
#   паразитную колонку col[2] = 264 DXA (~5 мм, меньше двойного cell-margin).
#   Word вынужден был держать 7 виртуальных слотов → все gridSpan во всех
#   строках сдвинулись (+1 у ячейки, перекрывавшей col[2]) → правые значения
#   схлопывались/уезжали. apply_layout читает span из документа, поэтому
#   сам по себе не мог это починить.
#
#   Решение: resolve_layout_template определяет режим по сумме span в строках.
#   Если сумма span = 7 — документ «отравлен» ghost-колонкой, apply_layout
#   вызывается с fix_ghost_spans=True.

ROW_HEIGHT = "340"
ROW_HEIGHT_RULE = "atLeast"

GRID_COLS_6 = ["2041", "1757", "1787", "1898", "1701", "1646"]  # базовый шаблон
GRID_COLS_7 = ["2000", "1798", "1787", "1898", "1701", "1646"]  # 7-кол документы (тоже 6 чисел!)
# Обе суммы = 10830. Проверка: assert sum(int(w) for w in GRID_COLS_6) == 10830

TABLE_WIDTH = "10830"  # жёстко — не вычислять из ширины страницы

# Позиция ghost-колонки в сломанных документах (0-based)
_GHOST_COL_IDX = 2
DEFAULT_GRID_COLS = GRID_COLS_6


def resolve_layout_template(template_name: str = "default") -> list[str]:
    """
    Возвращает список ширин колонок (DXA) по имени шаблона.
    Совместимость: 'default'/'6col' → GRID_COLS_6, '7col' → GRID_COLS_7.
    """
    if template_name == "7col":
        return GRID_COLS_7
    return GRID_COLS_6  # 'default', '6col', None, любой неизвестный


def hardcoded_layout(template_name: str = "default") -> list[str]:
    """Возвращает legacy-словарь, который ждут webapp/main.py и скрипты."""
    cols = list(resolve_layout_template(template_name))
    return {
        "template": "hardcoded",
        "grid_cols": cols,
        "grid_cols_6": list(GRID_COLS_6),
        "grid_cols_7": list(GRID_COLS_7),
        "tblGrid": None,
    }


def _build_cumsum(cols: list[str]) -> list[int]:
    cs = [0]
    for w in cols:
        cs.append(cs[-1] + int(w))
    return cs


def _build_tblGrid(cols: list[str]) -> etree._Element:
    """Строит элемент tblGrid из списка ширин."""
    tblGrid = etree.Element(qn("w:tblGrid"))
    for w in cols:
        col = etree.SubElement(tblGrid, qn("w:gridCol"))
        col.set(qn("w:w"), w)
    return tblGrid


def _fix_ghost_spans(spans: list[int]) -> list[int]:
    """
    Убирает влияние ghost-колонки на позиции _GHOST_COL_IDX.

    Когда документ был сгенерирован с 7 gridCol (ghost col[2]=264),
    ровно одна ячейка в каждой строке получила span+1 — та, чья зона
    покрывала позицию _GHOST_COL_IDX.
    Функция находит эту ячейку и уменьшает её span обратно на 1.

    Проверено на 24 строках реального документа: 24/24 ✓.
    """
    fixed = []
    col_idx = 0
    for span in spans:
        if col_idx <= _GHOST_COL_IDX < col_idx + span:
            fixed.append(max(1, span - 1))
        else:
            fixed.append(span)
        col_idx += span
    return fixed


def _detect_ghost_cols(table) -> bool:
    """
    Возвращает True, если таблица содержит ghost-колонку:
    сумма span хотя бы в одной строке равна 7 (вместо 6).
    """
    for row in table.rows:
        spans_sum = 0
        for tc in row._tr.findall(qn("w:tc")):
            tcPr = tc.find(qn("w:tcPr"))
            gs = tcPr.find(qn("w:gridSpan")) if tcPr is not None else None
            spans_sum += int(gs.get(qn("w:val"))) if gs is not None else 1
        if spans_sum == 7:
            return True
    return False


def _main_table_indices(doc) -> list[int]:
    """Берём самую большую таблицу отчёта, fallback — все >=3 строк."""
    if not doc.tables:
        return []
    scored = [(i, len(t.rows)) for i, t in enumerate(doc.tables)]
    best_i, best_n = max(scored, key=lambda x: x[1])
    if best_n >= 8:
        return [best_i]
    return [i for i, n in scored if n >= 3] or [0]


def diagnose_document(doc, layout: dict | None = None) -> list[str]:
    """Минимальная диагностика геометрии таблиц для /diagnose/reports."""
    grid_cols = (layout or {}).get("grid_cols") or list(DEFAULT_GRID_COLS)
    expected = len(grid_cols)
    out: list[str] = []
    if not doc.tables:
        return ["нет таблиц"]
    for i, table in enumerate(doc.tables):
        issues: list[str] = []
        grid = table._tbl.find(qn("w:tblGrid"))
        if grid is not None:
            actual_cols = len(grid.findall(qn("w:gridCol")))
            if actual_cols != expected:
                issues.append(f"в файле {actual_cols} колонок сетки, ожид. {expected}")
        if _detect_ghost_cols(table):
            issues.append("обнаружены строки с sum(span)=7 (ghost)")
        if issues:
            out.append(f"табл.{i + 1} ({len(table.rows)} стр.): " + "; ".join(issues))
    return out


def apply_layout(
    doc,
    layout: dict = None,
    only_main_table: bool = False,
    fix_ghost_spans: bool = False,
    cols: list[str] | None = None,
):
    """
    Применяет к каждой таблице документа:
    - общую ширину таблицы (tblW = 10830, жёстко)
    - фиксированную сетку столбцов (tblGrid из cols или GRID_COLS_6)
    - ширину каждой ячейки по её gridSpan
    - фиксированную высоту каждой строки

    Параметры:
        layout:           словарь layout (legacy), может содержать "grid_cols"
        only_main_table:  если True — применить только к основной таблице отчёта
        fix_ghost_spans:  если True — исправить span перед расчётом ширин ячеек
        cols:             явный список ширин; если None — автовыбор по документу
    """
    warnings: list[str] = []
    if cols is None:
        if isinstance(layout, dict) and layout.get("grid_cols"):
            cols = list(layout["grid_cols"])
        else:
            cols = GRID_COLS_6  # дефолт; вызывающий код может передать GRID_COLS_7

    cumsum = _build_cumsum(cols)
    indices = _main_table_indices(doc) if only_main_table else list(range(len(doc.tables)))

    for i in indices:
        table = doc.tables[i]
        tbl = table._tbl

        # Автоопределение ghost-режима, если не задано явно
        needs_fix = fix_ghost_spans or _detect_ghost_cols(table)
        if needs_fix:
            warnings.append(f"табл.{i + 1}: обнаружена ghost-колонка (sum_span=7), spans скорректированы")

        # ── tblPr: ширина таблицы и запрет автоподбора ────────────────────────
        tblPr = tbl.find(qn("w:tblPr"))
        if tblPr is None:
            tblPr = etree.SubElement(tbl, qn("w:tblPr"))
            tbl.insert(0, tblPr)

        tblW = tblPr.find(qn("w:tblW"))
        if tblW is None:
            tblW = etree.SubElement(tblPr, qn("w:tblW"))
        tblW.set(qn("w:w"), TABLE_WIDTH)  # ← жёстко 10830
        tblW.set(qn("w:type"), "dxa")

        tblLayout = tblPr.find(qn("w:tblLayout"))
        if tblLayout is None:
            tblLayout = etree.SubElement(tblPr, qn("w:tblLayout"))
        tblLayout.set(qn("w:type"), "fixed")

        # ── tblGrid ────────────────────────────────────────────────────────────
        old_grid = tbl.find(qn("w:tblGrid"))
        new_grid = _build_tblGrid(cols)
        if old_grid is not None:
            tbl.replace(old_grid, new_grid)
        else:
            tblPr.addnext(new_grid)

        # ── Строки: высота + ширины ячеек ─────────────────────────────────────
        for row in table.rows:
            tr = row._tr

            # Высота строки
            trPr = tr.find(qn("w:trPr"))
            if trPr is None:
                trPr = etree.SubElement(tr, qn("w:trPr"))
                tr.insert(0, trPr)
            trHeight = trPr.find(qn("w:trHeight"))
            if trHeight is None:
                trHeight = etree.SubElement(trPr, qn("w:trHeight"))
            trHeight.set(qn("w:val"), ROW_HEIGHT)
            trHeight.set(qn("w:hRule"), ROW_HEIGHT_RULE)

            # Собираем spans текущей строки
            tcs = tr.findall(qn("w:tc"))
            raw_spans = []
            for tc in tcs:
                tcPr = tc.find(qn("w:tcPr"))
                gs = tcPr.find(qn("w:gridSpan")) if tcPr is not None else None
                raw_spans.append(int(gs.get(qn("w:val"))) if gs is not None else 1)

            # Исправляем ghost-spans при необходимости
            spans = _fix_ghost_spans(raw_spans) if needs_fix else raw_spans

            # Применяем скорректированные spans и считаем ширины
            col_idx = 0
            for tc, span in zip(tcs, spans):
                if col_idx >= len(cols):
                    break

                span = max(1, min(span, len(cols) - col_idx))
                cell_w = str(cumsum[col_idx + span] - cumsum[col_idx])

                tcPr = tc.find(qn("w:tcPr"))
                if tcPr is None:
                    tcPr = etree.SubElement(tc, qn("w:tcPr"))
                    tc.insert(0, tcPr)

                # Обновляем gridSpan (важно: записываем исправленный span!)
                gs_el = tcPr.find(qn("w:gridSpan"))
                if span > 1:
                    if gs_el is None:
                        gs_el = etree.SubElement(tcPr, qn("w:gridSpan"))
                    gs_el.set(qn("w:val"), str(span))
                elif gs_el is not None:
                    # span=1 → элемент gridSpan не нужен
                    tcPr.remove(gs_el)

                tcW = tcPr.find(qn("w:tcW"))
                if tcW is None:
                    tcW = etree.SubElement(tcPr, qn("w:tcW"))
                tcW.set(qn("w:w"), cell_w)
                tcW.set(qn("w:type"), "dxa")

                col_idx += span
    return warnings


def read_template_layout(template_path: Path) -> dict:
    """Читает tblGrid из шаблона. Устарело — используется только для совместимости."""
    doc = Document(os.fspath(template_path))
    tbl = doc.tables[0]._tbl
    tblGrid = tbl.find(qn("w:tblGrid"))
    grid_cols = []
    if tblGrid is not None:
        for col in tblGrid.findall(qn("w:gridCol")):
            w = col.get(qn("w:w"))
            if w:
                grid_cols.append(w)
    return {
        "template": str(template_path),
        "tblGrid": deepcopy(tblGrid) if tblGrid is not None else None,
        "grid_cols": grid_cols or list(DEFAULT_GRID_COLS),
    }


def main():
    if len(sys.argv) < 2:
        print("Использование: python3 apply_template_layout.py <document.docx>")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    if not input_path.exists():
        print(f"Файл не найден: {input_path}")
        sys.exit(1)

    doc = Document(os.fspath(input_path))
    print(f"Документ: {len(doc.tables)} таблиц, строк: {[len(t.rows) for t in doc.tables]}")

    # Автоопределение: если в таблице есть строки с sum(span)=7 — ghost-режим
    has_ghost = any(_detect_ghost_cols(t) for t in doc.tables)
    cols = GRID_COLS_7 if has_ghost else GRID_COLS_6
    if has_ghost:
        print("Обнаружена ghost-колонка (sum_span=7) — применяем коррекцию spans")

    apply_layout(doc, cols=cols)

    output_path = input_path.parent / f"{input_path.stem}_layout.docx"
    doc.save(os.fspath(output_path))
    print(f"Сохранён: {output_path}")


if __name__ == "__main__":
    main()
