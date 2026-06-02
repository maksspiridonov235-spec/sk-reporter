"""
Применяет табличную разметку к отчётам:
- 6-колоночные таблицы остаются 6-колоночными (эталонный шаблон),
- 7-колоночные (Громов) остаются 7-колоночными.
"""

import os
import sys
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

TEMPLATE_PATH = (
    Path(__file__).parent
    / "contractor_report"
    / "болванки (шаблоны не вырезать только копировать)"
    / "Ежедневный отчет Шаблон.docx"
)

ROW_HEIGHT = "340"
ROW_HEIGHT_RULE = "atLeast"
GRID_COLS_6 = ["2041", "1757", "1787", "1898", "1701", "1646"]
GRID_COLS_7 = ["2041", "1757", "1787", "1898", "1701", "1291", "355"]


MIN_ROWS_FOR_MAIN_TABLE = 8


def _main_table_indices(doc) -> list[int]:
    """Сетку на таблицы-отчёты (>= MIN_ROWS), в т.ч. в merged.docx."""
    if not doc.tables:
        return []
    scored = [(i, len(t.rows)) for i, t in enumerate(doc.tables)]
    qualifying = [i for i, n in scored if n >= MIN_ROWS_FOR_MAIN_TABLE]
    if qualifying:
        return qualifying
    return [i for i, n in scored if n >= 3] or [0]


def _set_tc_width(tc, width_dxa: str) -> None:
    tc_pr = tc.find(qn("w:tcPr"))
    if tc_pr is None:
        tc_pr = etree.SubElement(tc, qn("w:tcPr"))
        tc.insert(0, tc_pr)
    tc_w = tc_pr.find(qn("w:tcW"))
    if tc_w is None:
        tc_w = etree.SubElement(tc_pr, qn("w:tcW"))
    tc_w.set(qn("w:w"), width_dxa)
    tc_w.set(qn("w:type"), "dxa")


def _build_cumsum(grid_cols: list[str]) -> list[int]:
    cumsum = [0]
    for width in grid_cols:
        cumsum.append(cumsum[-1] + int(width))
    return cumsum


def resolve_layout_template(templates_dir: Path | None = None) -> Path:
    """Возвращает путь к шаблону для layout."""
    if templates_dir is None:
        return TEMPLATE_PATH
    templates_dir = Path(templates_dir)
    direct = templates_dir / "Ежедневный отчет Шаблон.docx"
    if direct.exists():
        return direct
    candidates = sorted(templates_dir.glob("*Шаблон*.docx"))
    if candidates:
        return candidates[0]
    raise FileNotFoundError(f"Шаблон layout не найден в {templates_dir}")


def hardcoded_layout() -> dict:
    """Совместимость со старым импортом в webapp/main.py."""
    return {
        "template": "hardcoded",
        "grid_cols": list(GRID_COLS_6),
        "grid_cols_6": list(GRID_COLS_6),
        "grid_cols_7": list(GRID_COLS_7),
        "tblGrid": None,
    }


def read_template_layout(template_path: Path) -> dict:
    """Читает tblGrid из шаблона (совместимость API)."""
    doc = Document(os.fspath(template_path))
    tbl = doc.tables[0]._tbl if doc.tables else None
    tbl_grid = tbl.find(qn("w:tblGrid")) if tbl is not None else None
    return {"tblGrid": deepcopy(tbl_grid) if tbl_grid is not None else None}


def _build_tbl_grid(grid_cols: list[str]) -> etree._Element:
    tbl_grid = etree.Element(qn("w:tblGrid"))
    for width in grid_cols:
        col = etree.SubElement(tbl_grid, qn("w:gridCol"))
        col.set(qn("w:w"), width)
    return tbl_grid


def _detect_table_column_count(tbl) -> int:
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        grid_cols = grid.findall(qn("w:gridCol"))
        if len(grid_cols) in (6, 7):
            return len(grid_cols)

    max_span = 0
    for tr in tbl.findall(qn("w:tr")):
        span_sum = 0
        for tc in tr.findall(qn("w:tc")):
            tc_pr = tc.find(qn("w:tcPr"))
            gs_el = tc_pr.find(qn("w:gridSpan")) if tc_pr is not None else None
            span_sum += int(gs_el.get(qn("w:val"))) if gs_el is not None else 1
        max_span = max(max_span, span_sum)
    return 7 if max_span == 7 else 6


def apply_layout(doc, layout: dict | None = None, only_main_table: bool = True) -> list[str]:
    """
    Применяет layout к таблицам отчёта:
    - 6-колонок -> сетка 6,
    - 7-колонок -> сетка 7.
    """
    _ = layout
    indices = _main_table_indices(doc) if only_main_table else list(range(len(doc.tables)))

    for ti in indices:
        table = doc.tables[ti]
        tbl = table._tbl
        col_count = _detect_table_column_count(tbl)
        grid_cols = GRID_COLS_7 if col_count == 7 else GRID_COLS_6
        grid_cumsum = _build_cumsum(grid_cols)
        table_width = str(sum(int(w) for w in grid_cols))

        tbl_pr = tbl.find(qn("w:tblPr"))
        if tbl_pr is None:
            tbl_pr = etree.SubElement(tbl, qn("w:tblPr"))
            tbl.insert(0, tbl_pr)

        tbl_w = tbl_pr.find(qn("w:tblW"))
        if tbl_w is None:
            tbl_w = etree.SubElement(tbl_pr, qn("w:tblW"))
        tbl_w.set(qn("w:w"), table_width)
        tbl_w.set(qn("w:type"), "dxa")

        tbl_layout = tbl_pr.find(qn("w:tblLayout"))
        if tbl_layout is None:
            tbl_layout = etree.SubElement(tbl_pr, qn("w:tblLayout"))
        tbl_layout.set(qn("w:type"), "fixed")

        old_grid = tbl.find(qn("w:tblGrid"))
        new_grid = _build_tbl_grid(grid_cols)
        if old_grid is not None:
            tbl.replace(old_grid, new_grid)
        else:
            tbl_pr.addnext(new_grid)

        for row in table.rows:
            tr = row._tr
            col_idx = 0
            for tc in tr.findall(qn("w:tc")):
                if col_idx >= len(grid_cols):
                    break

                tc_pr = tc.find(qn("w:tcPr"))
                span = 1
                if tc_pr is not None:
                    gs_el = tc_pr.find(qn("w:gridSpan"))
                    if gs_el is not None:
                        span = max(1, int(gs_el.get(qn("w:val"), 1)))
                span = min(span, len(grid_cols) - col_idx)
                cell_w = str(grid_cumsum[col_idx + span] - grid_cumsum[col_idx])
                _set_tc_width(tc, cell_w)
                col_idx += span

    return []


def main():
    if len(sys.argv) < 2:
        print("Использование: python3 apply_template_layout.py <document.docx>")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    if not input_path.exists():
        print(f"Файл не найден: {input_path}")
        sys.exit(1)

    doc = Document(os.fspath(input_path))
    apply_layout(doc)
    output_path = input_path.parent / f"{input_path.stem}_layout.docx"
    doc.save(os.fspath(output_path))
    print(f"Сохранён: {output_path}")


if __name__ == "__main__":
    main()
