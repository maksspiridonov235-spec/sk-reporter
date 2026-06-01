"""
Применяет сетку столбцов из эталонного шаблона ко всем таблицам документа.
"""

import os
import sys
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

BOLVANKI_DIR = (
    Path(__file__).parent
    / "contractor_report"
    / "болванки (шаблоны не вырезать только копировать)"
)

# запасные значения, если в шаблоне нет tblGrid
DEFAULT_GRID_COLS = ["2041", "1757", "1787", "1898", "1701", "1646"]
ROW_HEIGHT = "340"
ROW_HEIGHT_RULE = "atLeast"


def resolve_layout_template(templates_dir: Path | None = None) -> Path:
    """Эталон для сетки: Шаблон.docx → *Шаблон* → ЮНС → любой docx в болванках."""
    templates_dir = Path(templates_dir or BOLVANKI_DIR)
    if not templates_dir.is_dir():
        raise FileNotFoundError(f"Папка болванок не найдена: {templates_dir}")

    for name in ("Ежедневный отчет Шаблон.docx", "Ежедневный отчёт Шаблон.docx"):
        p = templates_dir / name
        if p.is_file():
            return p

    for pattern in ("*Шаблон*.docx", "*шаблон*.docx"):
        found = sorted(templates_dir.glob(pattern))
        if found:
            return found[0]

    for pattern in ("ЮНС_*.docx", "Евракор_*.docx"):
        found = sorted(templates_dir.glob(pattern))
        if found:
            return found[0]

    any_docx = sorted(templates_dir.glob("*.docx"))
    if any_docx:
        return any_docx[0]

    raise FileNotFoundError(f"В {templates_dir} нет ни одного .docx для сетки таблицы")


def _grid_cols_from_tbl(tbl) -> list[str]:
    tblGrid = tbl.find(qn("w:tblGrid"))
    if tblGrid is None:
        return list(DEFAULT_GRID_COLS)
    cols = []
    for col in tblGrid.findall(qn("w:gridCol")):
        w = col.get(qn("w:w"))
        if w:
            cols.append(w)
    return cols or list(DEFAULT_GRID_COLS)


def read_template_layout(template_path: Path) -> dict:
    """Читает tblGrid и список ширин колонок из первой таблицы шаблона."""
    doc = Document(os.fspath(template_path))
    if not doc.tables:
        raise ValueError(f"В шаблоне нет таблиц: {template_path}")
    tbl = doc.tables[0]._tbl
    tblGrid = tbl.find(qn("w:tblGrid"))
    grid_cols = _grid_cols_from_tbl(tbl)
    return {
        "template": str(template_path),
        "tblGrid": deepcopy(tblGrid) if tblGrid is not None else None,
        "grid_cols": grid_cols,
    }


def _build_tblGrid(grid_cols: list[str]) -> etree._Element:
    tblGrid = etree.Element(qn("w:tblGrid"))
    for w in grid_cols:
        col = etree.SubElement(tblGrid, qn("w:gridCol"))
        col.set(qn("w:w"), w)
    return tblGrid


def _grid_cumsum(grid_cols: list[str]) -> list[int]:
    cs = [0]
    for w in grid_cols:
        cs.append(cs[-1] + int(w))
    return cs


def apply_layout(doc, layout: dict | None = None):
    """Фиксированная сетка столбцов (tblGrid + ширины ячеек по gridSpan)."""
    grid_cols = (layout or {}).get("grid_cols") or list(DEFAULT_GRID_COLS)
    cumsum = _grid_cumsum(grid_cols)
    table_width = str(cumsum[-1])

    for table in doc.tables:
        tbl = table._tbl

        tblPr = tbl.find(qn("w:tblPr"))
        if tblPr is None:
            tblPr = etree.SubElement(tbl, qn("w:tblPr"))
            tbl.insert(0, tblPr)

        tblW = tblPr.find(qn("w:tblW"))
        if tblW is None:
            tblW = etree.SubElement(tblPr, qn("w:tblW"))
        tblW.set(qn("w:w"), table_width)
        tblW.set(qn("w:type"), "dxa")

        tblLayout = tblPr.find(qn("w:tblLayout"))
        if tblLayout is None:
            tblLayout = etree.SubElement(tblPr, qn("w:tblLayout"))
        tblLayout.set(qn("w:type"), "fixed")

        old_grid = tbl.find(qn("w:tblGrid"))
        new_grid = _build_tblGrid(grid_cols)
        if old_grid is not None:
            tbl.replace(old_grid, new_grid)
        else:
            tblPr.addnext(new_grid)

        for row in table.rows:
            tr = row._tr
            col_idx = 0
            for tc in tr.findall(qn("w:tc")):
                if col_idx >= len(grid_cols):
                    break

                tcPr = tc.find(qn("w:tcPr"))
                if tcPr is None:
                    tcPr = etree.SubElement(tc, qn("w:tcPr"))
                    tc.insert(0, tcPr)

                v_merge = tcPr.find(qn("w:vMerge"))
                if v_merge is not None and v_merge.get(qn("w:val")) != "restart":
                    continue

                gs_el = tcPr.find(qn("w:gridSpan"))
                span = int(gs_el.get(qn("w:val"))) if gs_el is not None else 1
                span = max(1, min(span, len(grid_cols) - col_idx))

                cell_w = str(cumsum[col_idx + span] - cumsum[col_idx])
                tcW = tcPr.find(qn("w:tcW"))
                if tcW is None:
                    tcW = etree.SubElement(tcPr, qn("w:tcW"))
                tcW.set(qn("w:w"), cell_w)
                tcW.set(qn("w:type"), "dxa")

                col_idx += span


def main():
    if len(sys.argv) < 2:
        print("Использование: python3 apply_template_layout.py <document.docx> [эталон.docx]")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    template_path = Path(sys.argv[2]) if len(sys.argv) > 2 else resolve_layout_template()

    layout = read_template_layout(template_path)
    doc = Document(os.fspath(input_path))
    print(f"Эталон: {template_path.name}, колонки: {layout['grid_cols']}")
    apply_layout(doc, layout)

    output_path = input_path.parent / f"{input_path.stem}_layout.docx"
    doc.save(os.fspath(output_path))
    print(f"Сохранён: {output_path}")


if __name__ == "__main__":
    main()
