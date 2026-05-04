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
    - фиксированную сетку столбцов (tblGrid)
    - фиксированную высоту каждой строки
    """
    for table in doc.tables:
        tbl = table._tbl

        # Заменяем tblGrid
        old_grid = tbl.find(qn('w:tblGrid'))
        new_grid = _build_tblGrid()
        if old_grid is not None:
            tbl.replace(old_grid, new_grid)
        else:
            tblPr = tbl.find(qn('w:tblPr'))
            if tblPr is not None:
                tblPr.addnext(new_grid)
            else:
                tbl.insert(0, new_grid)

        # Устанавливаем высоту каждой строки
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
