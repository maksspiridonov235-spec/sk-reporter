"""Применяет ширины ячеек и высоты строк из шаблона к документу.
Использование: python3 apply_template_layout.py <document.docx>
"""

import sys
import os
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

TEMPLATE_PATH = Path(__file__).parent / "contractor_report" / "болванки (шаблоны не вырезать только копировать)" / "Ежедневный отчет Шаблон.docx"


def read_template_layout(template_path: Path) -> list:
    # Используем os.fspath для правильной обработки путей с Unicode на Windows
    doc = Document(os.fspath(template_path))
    layout = []
    for table in doc.tables:
        table_layout = []
        for row in table.rows:
            tr = row._tr
            trPr = tr.find(qn('w:trPr'))
            height_val = None
            height_rule = None
            if trPr is not None:
                trH = trPr.find(qn('w:trHeight'))
                if trH is not None:
                    height_val = trH.get(qn('w:val'))
                    height_rule = trH.get(qn('w:hRule'))

            cells_layout = []
            for tc in tr.findall(qn('w:tc')):
                tcPr = tc.find(qn('w:tcPr'))
                w_val = None
                w_type = 'dxa'
                grid_span = None
                if tcPr is not None:
                    tcW = tcPr.find(qn('w:tcW'))
                    if tcW is not None:
                        w_val = tcW.get(qn('w:w'))
                        w_type = tcW.get(qn('w:type')) or 'dxa'
                    gs = tcPr.find(qn('w:gridSpan'))
                    if gs is not None:
                        grid_span = gs.get(qn('w:val'))
                cells_layout.append({'w': w_val, 'type': w_type, 'gridSpan': grid_span})

            table_layout.append({
                'height': height_val,
                'hRule': height_rule,
                'cells': cells_layout,
            })
        layout.append(table_layout)
    return layout


def apply_layout(doc: Document, layout: list):
    for ti, table in enumerate(doc.tables):
        if ti >= len(layout):
            break
        table_layout = layout[ti]
        for ri, row in enumerate(table.rows):
            if ri >= len(table_layout):
                break
            row_layout = table_layout[ri]

            tr = row._tr
            trPr = tr.find(qn('w:trPr'))
            if trPr is None:
                trPr = etree.SubElement(tr, qn('w:trPr'))
                tr.insert(0, trPr)

            if row_layout['height'] is not None:
                trH = trPr.find(qn('w:trHeight'))
                if trH is None:
                    trH = etree.SubElement(trPr, qn('w:trHeight'))
                trH.set(qn('w:val'), row_layout['height'])
                if row_layout['hRule']:
                    trH.set(qn('w:hRule'), row_layout['hRule'])
                else:
                    if qn('w:hRule') in trH.attrib:
                        del trH.attrib[qn('w:hRule')]

            tcs = tr.findall(qn('w:tc'))
            cells_layout = row_layout['cells']
            if len(tcs) != len(cells_layout):
                continue

            for tc, cl in zip(tcs, cells_layout):
                if cl['w'] is None:
                    continue
                tcPr = tc.find(qn('w:tcPr'))
                if tcPr is None:
                    tcPr = etree.SubElement(tc, qn('w:tcPr'))
                    tc.insert(0, tcPr)

                tcW = tcPr.find(qn('w:tcW'))
                if tcW is None:
                    tcW = etree.SubElement(tcPr, qn('w:tcW'))
                tcW.set(qn('w:w'), cl['w'])
                tcW.set(qn('w:type'), cl['type'])

                gs_el = tcPr.find(qn('w:gridSpan'))
                if cl['gridSpan'] is not None:
                    if gs_el is None:
                        gs_el = etree.SubElement(tcPr, qn('w:gridSpan'))
                    gs_el.set(qn('w:val'), cl['gridSpan'])
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

    layout = read_template_layout(TEMPLATE_PATH)
    print(f"Шаблон: {len(layout)} таблиц")

    doc = Document(os.fspath(input_path))
    print(f"Документ: {len(doc.tables)} таблиц")

    apply_layout(doc, layout)

    output_path = input_path.parent / f"{input_path.stem}_layout.docx"
    doc.save(os.fspath(output_path))
    print(f"Сохранён: {output_path}")


if __name__ == "__main__":
    main()
