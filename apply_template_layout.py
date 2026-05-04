"""
Применяет ширины ячеек, высоты строк и сетку столбцов из шаблона к документу.
Сопоставление — по сигнатуре строки (кол-во ячеек + gridSpan),
поэтому работает с документами с любым количеством строк.

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


def _row_signature(tr) -> tuple:
    tcs = tr.findall(qn('w:tc'))
    spans = []
    for tc in tcs:
        tcPr = tc.find(qn('w:tcPr'))
        gs = None
        if tcPr is not None:
            gs_el = tcPr.find(qn('w:gridSpan'))
            if gs_el is not None:
                gs = gs_el.get(qn('w:val'))
        spans.append(gs or '1')
    return (len(tcs), tuple(spans))


def _row_cell_layout(tr) -> list:
    cells = []
    for tc in tr.findall(qn('w:tc')):
        tcPr = tc.find(qn('w:tcPr'))
        w_val, w_type, grid_span = None, 'dxa', None
        if tcPr is not None:
            tcW = tcPr.find(qn('w:tcW'))
            if tcW is not None:
                w_val = tcW.get(qn('w:w'))
                w_type = tcW.get(qn('w:type')) or 'dxa'
            gs_el = tcPr.find(qn('w:gridSpan'))
            if gs_el is not None:
                grid_span = gs_el.get(qn('w:val'))
        cells.append({'w': w_val, 'type': w_type, 'gridSpan': grid_span})
    return cells


def _row_height_layout(tr) -> dict:
    trPr = tr.find(qn('w:trPr'))
    height_val, height_rule = None, None
    if trPr is not None:
        trH = trPr.find(qn('w:trHeight'))
        if trH is not None:
            height_val = trH.get(qn('w:val'))
            height_rule = trH.get(qn('w:hRule'))
    return {'height': height_val, 'hRule': height_rule}


def read_template_index(template_path: Path) -> list:
    doc = Document(os.fspath(template_path))
    tables_index = []

    for table in doc.tables:
        tbl = table._tbl
        tblGrid = tbl.find(qn('w:tblGrid'))
        grid_copy = deepcopy(tblGrid) if tblGrid is not None else None

        sig_to_layout = {}
        ordered = []

        for row in table.rows:
            tr = row._tr
            sig = _row_signature(tr)
            if sig not in sig_to_layout:
                sig_to_layout[sig] = {
                    **_row_height_layout(tr),
                    'cells': _row_cell_layout(tr),
                    'sig': sig,
                }
                ordered.append(sig)

        tables_index.append({
            'map': sig_to_layout,
            'order': ordered,
            'tblGrid': grid_copy,
        })

    return tables_index


def _apply_row_layout(tr, layout: dict):
    trPr = tr.find(qn('w:trPr'))
    if trPr is None:
        trPr = etree.SubElement(tr, qn('w:trPr'))
        tr.insert(0, trPr)

    if layout['height'] is not None:
        trH = trPr.find(qn('w:trHeight'))
        if trH is None:
            trH = etree.SubElement(trPr, qn('w:trHeight'))
        trH.set(qn('w:val'), layout['height'])
        if layout['hRule']:
            trH.set(qn('w:hRule'), layout['hRule'])
        else:
            trH.attrib.pop(qn('w:hRule'), None)

    tcs = tr.findall(qn('w:tc'))
    cells_layout = layout['cells']
    if len(tcs) != len(cells_layout):
        return

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


def apply_layout(doc, tables_index: list):
    unmatched = 0

    for ti, table in enumerate(doc.tables):
        if ti >= len(tables_index):
            break
        idx = tables_index[ti]
        sig_map = idx['map']

        # Заменяем tblGrid из шаблона
        if idx['tblGrid'] is not None:
            tbl = table._tbl
            old_grid = tbl.find(qn('w:tblGrid'))
            new_grid = deepcopy(idx['tblGrid'])
            if old_grid is not None:
                tbl.replace(old_grid, new_grid)
            else:
                # Вставляем после tblPr
                tblPr = tbl.find(qn('w:tblPr'))
                if tblPr is not None:
                    tblPr.addnext(new_grid)
                else:
                    tbl.insert(0, new_grid)

        for row in table.rows:
            tr = row._tr
            sig = _row_signature(tr)

            if sig in sig_map:
                _apply_row_layout(tr, sig_map[sig])
            else:
                cell_count = sig[0]
                fallback = next(
                    (sig_map[s] for s in idx['order'] if s[0] == cell_count),
                    None
                )
                if fallback:
                    _apply_row_layout(tr, fallback)
                else:
                    unmatched += 1
                    print(f"  [!] Таблица {ti}, строка без образца: sig={sig}")

    if unmatched:
        print(f"Итого строк без образца: {unmatched}")


def main():
    if len(sys.argv) < 2:
        print("Использование: python3 apply_template_layout.py <document.docx>")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    if not input_path.exists():
        print(f"Файл не найден: {input_path}")
        sys.exit(1)

    print(f"Читаю шаблон: {TEMPLATE_PATH}")
    tables_index = read_template_index(TEMPLATE_PATH)
    print(f"  Шаблон: {len(tables_index)} таблиц, "
          f"уникальных сигнатур: {[len(t['map']) for t in tables_index]}")

    doc = Document(os.fspath(input_path))
    print(f"  Документ: {len(doc.tables)} таблиц, "
          f"строк: {[len(t.rows) for t in doc.tables]}")

    apply_layout(doc, tables_index)

    output_path = input_path.parent / f"{input_path.stem}_layout.docx"
    doc.save(os.fspath(output_path))
    print(f"Сохранён: {output_path}")


if __name__ == "__main__":
    main()
