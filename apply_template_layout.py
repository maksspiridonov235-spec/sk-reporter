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

# Фиксированные значения из шаблона (сумма каждого набора = 10830 DXA)
ROW_HEIGHT = "340"
ROW_HEIGHT_RULE = "atLeast"
GRID_COLS_6 = ["2041", "1757", "1787", "1898", "1701", "1646"]
GRID_COLS_7 = ["2000", "1798", "1787", "1898", "1701", "1646"]
GRID_COLS = GRID_COLS_6
DEFAULT_GRID_COLS = GRID_COLS_6
TABLE_WIDTH = str(sum(int(w) for w in GRID_COLS))  # 10830
MIN_ROWS_FOR_MAIN_TABLE = 8


def _grid_cumsum(grid_cols: list[str]) -> list[int]:
    cs = [0]
    for w in grid_cols:
        cs.append(cs[-1] + int(w))
    return cs


def hardcoded_layout() -> dict:
    """Сетка по умолчанию для webapp и prepare-пайплайна."""
    return {
        "template": "hardcoded",
        "grid_cols": list(GRID_COLS_6),
        "grid_cols_6": list(GRID_COLS_6),
        "grid_cols_7": list(GRID_COLS_7),
        "tblGrid": None,
    }


def resolve_layout_template(templates_dir: Path | None = None) -> Path:
    """Путь к .docx-шаблону с эталонной сеткой таблицы."""
    if templates_dir is None:
        return TEMPLATE_PATH
    templates_dir = Path(templates_dir)
    direct = templates_dir / "Ежедневный отчет Шаблон.docx"
    if direct.is_file():
        return direct
    candidates = sorted(templates_dir.glob("*Шаблон*.docx"))
    if candidates:
        return candidates[0]
    raise FileNotFoundError(f"Шаблон layout не найден в {templates_dir}")


def read_template_layout(template_path: Path) -> dict:
    """Читает tblGrid из шаблона для применения к документам."""
    doc = Document(os.fspath(template_path))
    tbl = doc.tables[0]._tbl
    tblGrid = tbl.find(qn('w:tblGrid'))
    grid_cols = list(GRID_COLS)
    if tblGrid is not None:
        cols = []
        for col in tblGrid.findall(qn('w:gridCol')):
            w = col.get(qn('w:w'))
            if w:
                cols.append(w)
        if cols:
            grid_cols = cols
    return {
        'template': str(template_path),
        'tblGrid': deepcopy(tblGrid) if tblGrid is not None else None,
        'grid_cols': grid_cols,
    }


def _build_tblGrid(grid_cols: list[str] | None = None) -> etree._Element:
    """Строит элемент tblGrid из списка ширин колонок."""
    cols = grid_cols or GRID_COLS
    tblGrid = etree.Element(qn('w:tblGrid'))
    for w in cols:
        col = etree.SubElement(tblGrid, qn('w:gridCol'))
        col.set(qn('w:w'), w)
    return tblGrid


def _row_start_col(tr) -> int:
    tr_pr = tr.find(qn('w:trPr'))
    if tr_pr is None:
        return 0
    gb = tr_pr.find(qn('w:gridBefore'))
    if gb is None:
        return 0
    return int(gb.get(qn('w:val'), 0))


def _row_occupied_cols(tr, ncol: int) -> int:
    col_idx = _row_start_col(tr)
    for tc in tr.findall(qn('w:tc')):
        if col_idx >= ncol:
            break
        tc_pr = tc.find(qn('w:tcPr'))
        span = 1
        if tc_pr is not None:
            gs = tc_pr.find(qn('w:gridSpan'))
            if gs is not None:
                span = max(1, int(gs.get(qn('w:val'), 1)))
            vm = tc_pr.find(qn('w:vMerge'))
            if vm is not None and vm.get(qn('w:val')) != 'restart':
                col_idx += span
                continue
        col_idx += span
    return col_idx


def diagnose_table(tbl, grid_cols: list[str]) -> list[str]:
    ncol = len(grid_cols)
    issues: list[str] = []
    rows = tbl.findall(qn('w:tr'))
    tbl_grid = tbl.find(qn('w:tblGrid'))
    file_ncol = len(tbl_grid.findall(qn('w:gridCol'))) if tbl_grid is not None else 0
    if file_ncol and file_ncol != ncol:
        issues.append(f"в файле {file_ncol} колонок сетки, ожид. {ncol}")
    bad = []
    for ri, tr in enumerate(rows):
        occ = _row_occupied_cols(tr, ncol)
        if occ != ncol:
            bad.append(f"{ri + 1}({occ})")
    if bad:
        preview = ", ".join(bad[:6])
        more = f" +{len(bad) - 6}" if len(bad) > 6 else ""
        issues.append(f"битые строки: {preview}{more}")
    return issues


def diagnose_document(doc, layout: dict | None = None) -> list[str]:
    grid_cols = (layout or {}).get('grid_cols') or list(DEFAULT_GRID_COLS)
    out: list[str] = []
    if not doc.tables:
        out.append('нет таблиц')
        return out
    for i, table in enumerate(doc.tables):
        issues = diagnose_table(table._tbl, grid_cols)
        if issues:
            out.append(f"табл.{i + 1} ({len(table.rows)} стр.): " + "; ".join(issues))
    return out


def _main_table_indices(doc) -> list[int]:
    if not doc.tables:
        return []
    scored = [(i, len(t.rows)) for i, t in enumerate(doc.tables)]
    best_i, best_n = max(scored, key=lambda x: x[1])
    if best_n >= MIN_ROWS_FOR_MAIN_TABLE:
        return [best_i]
    return [i for i, n in scored if n >= 3] or [0]


def apply_layout(doc, layout: dict = None, only_main_table: bool = False):
    """
    Применяет к каждой таблице документа:
    - общую ширину таблицы (tblW)
    - фиксированную сетку столбцов (tblGrid)
    - ширину каждой ячейки по её gridSpan
    - фиксированную высоту каждой строки
    - обнуление отступов в пустых ячейках

    Возвращает список предупреждений по сетке (диагностика до правок).
    """
    grid_cols = (layout or {}).get('grid_cols') or list(GRID_COLS)
    cumsum = _grid_cumsum(grid_cols)
    table_width = str(cumsum[-1])
    warnings = diagnose_document(doc, layout or {'grid_cols': grid_cols})
    indices = _main_table_indices(doc) if only_main_table else list(range(len(doc.tables)))

    for ti in indices:
        table = doc.tables[ti]
        tbl = table._tbl

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

        old_grid = tbl.find(qn('w:tblGrid'))
        new_grid = _build_tblGrid(grid_cols)
        if old_grid is not None:
            tbl.replace(old_grid, new_grid)
        else:
            tblPr.addnext(new_grid)

        for row in table.rows:
            tr = row._tr
            col_idx = 0
            for tc in tr.findall(qn('w:tc')):
                if col_idx >= len(grid_cols):
                    break

                tcPr = tc.find(qn('w:tcPr'))
                if tcPr is None:
                    tcPr = etree.SubElement(tc, qn('w:tcPr'))
                    tc.insert(0, tcPr)

                gs_el = tcPr.find(qn('w:gridSpan'))
                span = int(gs_el.get(qn('w:val'))) if gs_el is not None else 1
                span = max(1, min(span, len(grid_cols) - col_idx))

                cell_w = str(cumsum[col_idx + span] - cumsum[col_idx])

                tcW = tcPr.find(qn('w:tcW'))
                if tcW is None:
                    tcW = etree.SubElement(tcPr, qn('w:tcW'))
                tcW.set(qn('w:w'), cell_w)
                tcW.set(qn('w:type'), 'dxa')

                col_idx += span
    return warnings


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
