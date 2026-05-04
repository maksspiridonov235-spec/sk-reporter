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
GRID_COLS = ["2041", "1757", "1787", "1898", "1701", "1646"]
TABLE_WIDTH = str(sum(int(w) for w in GRID_COLS))  # 10830

# Накопленные суммы для расчёта ширины ячейки по gridSpan
# GRID_CUMSUM[i] = сумма первых i колонок
_GRID_CUMSUM = [0]
for _w in GRID_COLS:
    _GRID_CUMSUM.append(_GRID_CUMSUM[-1] + int(_w))


def read_template_layout(template_path: Path) -> dict:
    """Читает tblGrid из шаблона для применения к документам."""
    doc = Document(os.fspath(template_path))
    tbl = doc.tables[0]._tbl
    tblGrid = tbl.find(qn('w:tblGrid'))
    return {'tblGrid': deepcopy(tblGrid) if tblGrid is not None else None}


def _build_tblGrid() -> etree._Element:
    """Строит элемент tblGrid из фиксированных значений."""
    tblGrid = etree.Element(qn('w:tblGrid'))
    for w in GRID_COLS:
        col = etree.SubElement(tblGrid, qn('w:gridCol'))
        col.set(qn('w:w'), w)
    return tblGrid


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

        # Ширина таблицы и запрет автоподбора
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = etree.SubElement(tbl, qn('w:tblPr'))
            tbl.insert(0, tblPr)

        tblW = tblPr.find(qn('w:tblW'))
        if tblW is None:
            tblW = etree.SubElement(tblPr, qn('w:tblW'))
        tblW.set(qn('w:w'), TABLE_WIDTH)
        tblW.set(qn('w:type'), 'dxa')

        tblLayout = tblPr.find(qn('w:tblLayout'))
        if tblLayout is None:
            tblLayout = etree.SubElement(tblPr, qn('w:tblLayout'))
        tblLayout.set(qn('w:type'), 'fixed')

        # Заменяем tblGrid
        old_grid = tbl.find(qn('w:tblGrid'))
        new_grid = _build_tblGrid()
        if old_grid is not None:
            tbl.replace(old_grid, new_grid)
        else:
            tblPr.addnext(new_grid)

        # Устанавливаем высоту каждой строки и обнуляем отступы в пустых ячейках
        for row in table.rows:
            tr = row._tr
            trPr = tr.find(qn('w:trPr'))
            if trPr is None:
                trPr = etree.Element(qn('w:trPr'))
                tr.insert(0, trPr)

            trH = trPr.find(qn('w:trHeight'))
            if trH is None:
                trH = etree.SubElement(trPr, qn('w:trHeight'))
            trH.set(qn('w:val'), ROW_HEIGHT)
            trH.set(qn('w:hRule'), ROW_HEIGHT_RULE)

            # Ширины ячеек по gridSpan + обнуляем отступы в пустых
            col_idx = 0
            for tc in tr.findall(qn('w:tc')):
                tcPr = tc.find(qn('w:tcPr'))
                if tcPr is None:
                    tcPr = etree.SubElement(tc, qn('w:tcPr'))
                    tc.insert(0, tcPr)

                gs_el = tcPr.find(qn('w:gridSpan'))
                span = int(gs_el.get(qn('w:val'))) if gs_el is not None else 1
                span = max(1, min(span, len(GRID_COLS) - col_idx))

                cell_w = str(_GRID_CUMSUM[col_idx + span] - _GRID_CUMSUM[col_idx])

                tcW = tcPr.find(qn('w:tcW'))
                if tcW is None:
                    tcW = etree.SubElement(tcPr, qn('w:tcW'))
                tcW.set(qn('w:w'), cell_w)
                tcW.set(qn('w:type'), 'dxa')

                col_idx += span

            # Обнуляем отступы параграфов в пустых ячейках
            for tc in tr.findall(qn('w:tc')):
                cell_text = ''.join(tc.itertext()).strip()
                if cell_text:
                    continue
                for p in tc.findall(qn('w:p')):
                    pPr = p.find(qn('w:pPr'))
                    if pPr is None:
                        pPr = etree.SubElement(p, qn('w:pPr'))
                    spacing = pPr.find(qn('w:spacing'))
                    if spacing is None:
                        spacing = etree.SubElement(pPr, qn('w:spacing'))
                    spacing.set(qn('w:before'), '0')
                    spacing.set(qn('w:after'), '0')
                    spacing.set(qn('w:line'), '240')
                    spacing.set(qn('w:lineRule'), 'auto')


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
