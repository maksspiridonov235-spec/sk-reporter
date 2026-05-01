"""
Применяет ширины ячеек и высоты строк шаблона к документу.
Использование: python3 apply_template_layout.py <document.docx>
"""

import sys
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

# Размеры захардкожены из «Ежедневный отчет Шаблон.docx»
# Все строки: height=343 twips, hRule=None (atLeast)
TEMPLATE_ROW_HEIGHT = "343"

TEMPLATE_CELL_WIDTHS = [
    ['5497', '5497', '5497', '1559', '1646', '1646'],
    ['2151', '1361', '1985', '4851', '4851', '4851'],
    ['2151', '1361', '1985', '4851', '4851', '4851'],
    ['7056', '7056', '7056', '7056', '3292', '3292'],
    ['2151', '4905', '4905', '4905', '1646', '1646'],
    ['2151', '4905', '4905', '4905', '1646', '1646'],
    ['2151', '1361', '1985', '1559', '1646', '1646'],
    ['2151', '1361', '1985', '1559', '1646', '1646'],
    ['2151', '4905', '4905', '4905', '1646', '1646'],
    ['10348', '10348', '10348', '10348', '10348', '10348'],
    ['3512', '3512', '3544', '3544', '3292', '3292'],
    ['2151', '8197', '8197', '8197', '8197', '8197'],
    ['2151', '8197', '8197', '8197', '8197', '8197'],
    ['2151', '8197', '8197', '8197', '8197', '8197'],
    ['2151', '8197', '8197', '8197', '8197', '8197'],
    ['2151', '3346', '3346', '3205', '3205', '1646'],
    ['2151', '3346', '3346', '3205', '3205', '1646'],
    ['3512', '3512', '1985', '1559', '1646', '1646'],
    ['2151', '4905', '4905', '4905', '3292', '3292'],
    ['2151', '4905', '4905', '4905', '3292', '3292'],
    ['10348', '10348', '10348', '10348', '10348', '10348'],
    ['7056', '7056', '7056', '7056', '1646', '1646'],
    ['7056', '7056', '7056', '7056', '1646', '1646'],
    ['7056', '7056', '7056', '7056', '1646', '1646'],
    ['3512', '3512', '3544', '3544', '1646', '1646'],
    ['3512', '3512', '3544', '3544', '1646', '1646'],
    ['10348', '10348', '10348', '10348', '10348', '10348'],
    ['7056', '7056', '7056', '7056', '1646', '1646'],
    ['7056', '7056', '7056', '7056', '1646', '1646'],
    ['3512', '3512', '3544', '3544', '1646', '1646'],
    ['10348', '10348', '10348', '10348', '10348', '10348'],
    ['10348', '10348', '10348', '10348', '10348', '10348'],
    ['3512', '3512', '3544', '3544', '1646', '1646'],
    ['10348', '10348', '10348', '10348', '10348', '10348'],
    ['10348', '10348', '10348', '10348', '10348', '10348'],
    ['3512', '3512', '3544', '3544', '1646', '1646'],
    ['3512', '3512', '6836', '6836', '6836', '6836'],
    ['3512', '3512', '6836', '6836', '6836', '6836'],
    ['3512', '3512', '6836', '6836', '6836', '6836'],
    ['3512', '3512', '3544', '3544', '3292', '3292'],
    ['3512', '3512', '3544', '3544', '3292', '3292'],
    ['3512', '3512', '3544', '3544', '3292', '3292'],
]


def apply_layout(doc: Document, layout=None):
    """Apply hardcoded template row heights and cell widths to the first table."""
    widths = TEMPLATE_CELL_WIDTHS
    if not doc.tables:
        return
    table = doc.tables[0]
    for ri, row in enumerate(table.rows):
        tr = row._tr
        trPr = tr.find(qn('w:trPr'))
        if trPr is None:
            trPr = etree.SubElement(tr, qn('w:trPr'))
            tr.insert(0, trPr)

        trH = trPr.find(qn('w:trHeight'))
        if trH is None:
            trH = etree.SubElement(trPr, qn('w:trHeight'))
        trH.set(qn('w:val'), TEMPLATE_ROW_HEIGHT)
        if qn('w:hRule') in trH.attrib:
            del trH.attrib[qn('w:hRule')]

        if ri >= len(widths):
            continue
        row_widths = widths[ri]
        for ci, cell in enumerate(row.cells):
            if ci >= len(row_widths):
                break
            tc = cell._tc
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is None:
                tcPr = etree.SubElement(tc, qn('w:tcPr'))
                tc.insert(0, tcPr)
            tcW = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcW = etree.SubElement(tcPr, qn('w:tcW'))
            tcW.set(qn('w:w'), row_widths[ci])
            tcW.set(qn('w:type'), 'dxa')


def main():
    if len(sys.argv) < 2:
        print("Использование: python3 apply_template_layout.py <document.docx>")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    if not input_path.exists():
        print(f"Файл не найден: {input_path}")
        sys.exit(1)

    doc = Document(str(input_path))
    print(f"Документ: {len(doc.tables)} таблиц")

    apply_layout(doc)

    output_path = input_path.parent / f"{input_path.stem}_layout.docx"
    doc.save(str(output_path))
    print(f"Сохранён: {output_path}")


if __name__ == "__main__":
    main()
