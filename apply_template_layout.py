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

# Фиксированные значения из шаблона
ROW_HEIGHT = "340"
ROW_HEIGHT_RULE = "atLeast"
GRID_COLS_6 = ["2041", "1757", "1787", "1898", "1701", "1646"]
# 7-колоночная сетка Громова: первые 5 как в 6-колонках, договор разбит на 1291+355
GRID_COLS_7 = ["2041", "1757", "1787", "1898", "1701", "1291", "355"]


def _build_cumsum(grid_cols: list[str]) -> list[int]:
    cumsum = [0]
    for width in grid_cols:
        cumsum.append(cumsum[-1] + int(width))
    return cumsum


def read_template_layout(template_path: Path) -> dict:
    """Читает tblGrid из шаблона для применения к документам."""
    doc = Document(os.fspath(template_path))
    tbl = doc.tables[0]._tbl
    tblGrid = tbl.find(qn('w:tblGrid'))
    return {'tblGrid': deepcopy(tblGrid) if tblGrid is not None else None}


def _build_tblGrid(grid_cols: list[str]) -> etree._Element:
    """Строит элемент tblGrid из фиксированных значений."""
    tblGrid = etree.Element(qn('w:tblGrid'))
    for w in grid_cols:
        col = etree.SubElement(tblGrid, qn('w:gridCol'))
        col.set(qn('w:w'), w)
    return tblGrid


def _detect_table_column_count(tbl) -> int:
    """Определяет логическое число колонок таблицы (6 или 7)."""
    grid = tbl.find(qn('w:tblGrid'))
    if grid is not None:
        grid_cols = grid.findall(qn('w:gridCol'))
        if len(grid_cols) in (6, 7):
            return len(grid_cols)

    max_span = 0
    for tr in tbl.findall(qn('w:tr')):
        span_sum = 0
        for tc in tr.findall(qn('w:tc')):
            tcPr = tc.find(qn('w:tcPr'))
            gs_el = tcPr.find(qn('w:gridSpan')) if tcPr is not None else None
            span_sum += int(gs_el.get(qn('w:val'))) if gs_el is not None else 1
        max_span = max(max_span, span_sum)
    return 7 if max_span == 7 else 6


def apply_layout(doc, layout: dict = None):
    """
    Применяет к каждой таблице документа:
    - общую ширину таблицы (tblW)
    - фиксированную сетку столбцов (tblGrid)
    - ширину каждой ячейки по её gridSpan
    - фиксированную высоту каждой строки
    - обнуление отступов в пустых ячейках
    """
    for table in doc.tables:
        tbl = table._tbl

        col_count = _detect_table_column_count(tbl)
        grid_cols = GRID_COLS_7 if col_count == 7 else GRID_COLS_6
        grid_cumsum = _build_cumsum(grid_cols)
        table_width = str(sum(int(w) for w in grid_cols))

        # Ширина таблицы и запрет автоподбора
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = etree.SubElement(tbl, qn('w:tblPr'))
            tbl.insert(0, tblPr)

        tblW = tblPr.find(qn('w:tblW'))
        if tblW is None:
            tblW = etree.SubElement(tblPr, qn('w:tblW'))
        tblW.set(qn('w:w'), table_width)
        tblW.set(qn('w:type'), 'dxa')

        tblLayout = tblPr.find(qn('w:tblLayout'))
        if tblLayout is None:
            tblLayout = etree.SubElement(tblPr, qn('w:tblLayout'))
        tblLayout.set(qn('w:type'), 'fixed')

        # Заменяем tblGrid
        old_grid = tbl.find(qn('w:tblGrid'))
        new_grid = _build_tblGrid(grid_cols)
        if old_grid is not None:
            tbl.replace(old_grid, new_grid)
        else:
            tblPr.addnext(new_grid)

        # Обрабатываем каждую строку — только ширины ячеек
        for row in table.rows:
            tr = row._tr
            tcs = tr.findall(qn('w:tc'))

            # Ширины ячеек по gridSpan
            col_idx = 0
            for tc in tcs:
                if col_idx >= len(grid_cols):
                    break

                tcPr = tc.find(qn('w:tcPr'))
                if tcPr is None:
                    tcPr = etree.SubElement(tc, qn('w:tcPr'))
                    tc.insert(0, tcPr)

                gs_el = tcPr.find(qn('w:gridSpan'))
                span = int(gs_el.get(qn('w:val'))) if gs_el is not None else 1
                span = max(1, min(span, len(grid_cols) - col_idx))

                cell_w = str(grid_cumsum[col_idx + span] - grid_cumsum[col_idx])

                tcW = tcPr.find(qn('w:tcW'))
                if tcW is None:
                    tcW = etree.SubElement(tcPr, qn('w:tcW'))
                tcW.set(qn('w:w'), cell_w)
                tcW.set(qn('w:type'), 'dxa')

                col_idx += span


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

    apply_layout(doc)

    output_path = input_path.parent / f"{input_path.stem}_layout.docx"
    doc.save(os.fspath(output_path))
    print(f"Сохранён: {output_path}")


if __name__ == "__main__":
    main()
