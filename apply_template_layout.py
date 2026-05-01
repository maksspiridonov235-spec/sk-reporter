"""
Применяет ширины ячеек и высоты строк шаблона к документу.
Использование: python3 apply_template_layout.py <document.docx>
"""

import sys
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

# Захардкожено из «Ежедневный отчет Шаблон.docx»
# Каждая строка — список реальных <w:tc>: (w, gridSpan или None)
# Все строки: height=343 twips, hRule=None (atLeast)
TEMPLATE_ROW_HEIGHT = "343"

TEMPLATE_ROWS = [
    # (w, gridSpan)
    [('5497', '3'), ('1559', None), ('1646', None), ('1646', None)],            # 0
    [('2151', None), ('1361', None), ('1985', None), ('4851', '3')],            # 1
    [('2151', None), ('1361', None), ('1985', None), ('4851', '3')],            # 2
    [('7056', '4'), ('3292', '2')],                                             # 3
    [('2151', None), ('4905', '3'), ('1646', None), ('1646', None)],            # 4
    [('2151', None), ('4905', '3'), ('1646', None), ('1646', None)],            # 5
    [('2151', None), ('1361', None), ('1985', None), ('1559', None), ('1646', None), ('1646', None)],  # 6
    [('2151', None), ('1361', None), ('1985', None), ('1559', None), ('1646', None), ('1646', None)],  # 7
    [('2151', None), ('4905', '3'), ('1646', None), ('1646', None)],            # 8
    [('10348', '6')],                                                           # 9
    [('3512', '2'), ('3544', '2'), ('3292', '2')],                              # 10
    [('2151', None), ('8197', '5')],                                            # 11
    [('2151', None), ('8197', '5')],                                            # 12
    [('2151', None), ('8197', '5')],                                            # 13
    [('2151', None), ('8197', '5')],                                            # 14
    [('2151', None), ('3346', '2'), ('3205', '2'), ('1646', None)],             # 15
    [('2151', None), ('3346', '2'), ('3205', '2'), ('1646', None)],             # 16
    [('3512', '2'), ('1985', None), ('1559', None), ('1646', None), ('1646', None)],  # 17
    [('2151', None), ('4905', '3'), ('3292', '2')],                             # 18
    [('2151', None), ('4905', '3'), ('3292', '2')],                             # 19
    [('10348', '6')],                                                           # 20
    [('7056', '4'), ('1646', None), ('1646', None)],                            # 21
    [('7056', '4'), ('1646', None), ('1646', None)],                            # 22
    [('7056', '4'), ('1646', None), ('1646', None)],                            # 23
    [('3512', '2'), ('3544', '2'), ('1646', None), ('1646', None)],             # 24
    [('3512', '2'), ('3544', '2'), ('1646', None), ('1646', None)],             # 25
    [('10348', '6')],                                                           # 26
    [('7056', '4'), ('1646', None), ('1646', None)],                            # 27
    [('7056', '4'), ('1646', None), ('1646', None)],                            # 28
    [('3512', '2'), ('3544', '2'), ('1646', None), ('1646', None)],             # 29
    [('10348', '6')],                                                           # 30
    [('10348', '6')],                                                           # 31
    [('3512', '2'), ('3544', '2'), ('1646', None), ('1646', None)],             # 32
    [('10348', '6')],                                                           # 33
    [('10348', '6')],                                                           # 34
    [('3512', '2'), ('3544', '2'), ('1646', None), ('1646', None)],             # 35
    [('3512', '2'), ('6836', '4')],                                             # 36
    [('3512', '2'), ('6836', '4')],                                             # 37
    [('3512', '2'), ('6836', '4')],                                             # 38
    [('3512', '2'), ('3544', '2'), ('3292', '2')],                              # 39
    [('3512', '2'), ('3544', '2'), ('3292', '2')],                              # 40
    [('3512', '2'), ('3544', '2'), ('3292', '2')],                              # 41
]


def apply_layout(doc: Document, layout=None):
    """Apply hardcoded template row heights and cell widths to the first table."""
    if not doc.tables:
        return
    table = doc.tables[0]
    for ri, row in enumerate(table.rows):
        tr = row._tr

        # Высота строки
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

        if ri >= len(TEMPLATE_ROWS):
            continue

        # Ширины ячеек — работаем с реальными <w:tc>, не с row.cells
        tcs = tr.findall(qn('w:tc'))
        tmpl_cells = TEMPLATE_ROWS[ri]
        if len(tcs) != len(tmpl_cells):
            continue  # структура строки не совпадает — не трогаем

        for tc, (w_val, grid_span) in zip(tcs, tmpl_cells):
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is None:
                tcPr = etree.SubElement(tc, qn('w:tcPr'))
                tc.insert(0, tcPr)

            tcW = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcW = etree.SubElement(tcPr, qn('w:tcW'))
            tcW.set(qn('w:w'), w_val)
            tcW.set(qn('w:type'), 'dxa')

            gs_el = tcPr.find(qn('w:gridSpan'))
            if grid_span is not None:
                if gs_el is None:
                    gs_el = etree.SubElement(tcPr, qn('w:gridSpan'))
                gs_el.set(qn('w:val'), grid_span)
            else:
                if gs_el is not None:
                    tcPr.remove(gs_el)


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
