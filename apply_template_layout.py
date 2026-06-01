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

# Захардкоженная сетка (как в «Ежедневный отчет Шаблон.docx»)
DEFAULT_GRID_COLS = ["2041", "1757", "1787", "1898", "1701", "1646"]
# Часть отчётов (в т.ч. Громов): 7-я узкая колонка
DEFAULT_GRID_COLS_7 = ["1974", "1698", "1725", "1837", "1647", "1594", "355"]
ROW_HEIGHT = "340"
ROW_HEIGHT_RULE = "atLeast"
MIN_ROW_HEIGHT_CM = 0.6
MIN_ROWS_FOR_MAIN_TABLE = 8


def hardcoded_layout() -> dict:
    return {
        "template": "hardcoded",
        "grid_cols": list(DEFAULT_GRID_COLS),
        "tblGrid": None,
    }


def resolve_layout_template(templates_dir: Path | None = None) -> Path:
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


def _row_start_col(tr) -> int:
    tr_pr = tr.find(qn("w:trPr"))
    if tr_pr is None:
        return 0
    gb = tr_pr.find(qn("w:gridBefore"))
    if gb is None:
        return 0
    return int(gb.get(qn("w:val"), 0))


def _row_occupied_cols(tr, ncol: int) -> int:
    col_idx = _row_start_col(tr)
    for tc in tr.findall(qn("w:tc")):
        if col_idx >= ncol:
            break
        tc_pr = tc.find(qn("w:tcPr"))
        span = 1
        if tc_pr is not None:
            gs = tc_pr.find(qn("w:gridSpan"))
            if gs is not None:
                span = max(1, int(gs.get(qn("w:val"), 1)))
            vm = tc_pr.find(qn("w:vMerge"))
            if vm is not None and vm.get(qn("w:val")) != "restart":
                col_idx += span
                continue
        col_idx += span
    return col_idx


def _normalize_row_tr(tr) -> None:
    """Убирает gridBefore — частый артефакт, ломает подсчёт колонок."""
    tr_pr = tr.find(qn("w:trPr"))
    if tr_pr is None:
        return
    gb = tr_pr.find(qn("w:gridBefore"))
    if gb is not None:
        tr_pr.remove(gb)


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



def _max_row_occupancy(tbl) -> int:
    max_occ = 0
    for tr in tbl.findall(qn("w:tr")):
        max_occ = max(max_occ, _row_occupied_cols(tr, 16))
    return max_occ


def _grid_cols_for_table(tbl, standard_6: list[str]) -> list[str]:
    """6 колонок → эталон; 7 колонок → сетка из файла (не ломаем Громова)."""
    file_cols = _grid_cols_from_tbl(tbl)
    max_occ = _max_row_occupancy(tbl)

    if len(file_cols) == 7 and max_occ == 7:
        return list(file_cols)
    if len(file_cols) == 6 and max_occ == 6:
        return list(standard_6)
    if max_occ == 7:
        return list(file_cols) if len(file_cols) == 7 else list(DEFAULT_GRID_COLS_7)
    if max_occ == 6:
        return list(standard_6)
    if file_cols:
        return file_cols
    return list(standard_6)


def diagnose_table(tbl, grid_cols: list[str]) -> list[str]:
    ncol = len(grid_cols)
    issues: list[str] = []
    rows = tbl.findall(qn("w:tr"))
    tbl_grid = tbl.find(qn("w:tblGrid"))
    file_ncol = len(tbl_grid.findall(qn("w:gridCol"))) if tbl_grid is not None else 0
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
    standard = (layout or {}).get("grid_cols") or list(DEFAULT_GRID_COLS)
    out: list[str] = []
    if not doc.tables:
        out.append("нет таблиц")
        return out
    for i, table in enumerate(doc.tables):
        tbl = table._tbl
        nrows = len(table.rows)
        expected = _grid_cols_for_table(tbl, standard)
        issues = diagnose_table(tbl, expected)
        if issues:
            out.append(f"табл.{i + 1} ({nrows} стр.): " + "; ".join(issues))
        elif len(expected) == 7:
            out.append(f"табл.{i + 1} ({nrows} стр.): 7 колонок — своя сетка (OK)")
    return out


def _main_table_indices(doc) -> list[int]:
    """Сетку вешаем на большую таблицу отчёта, мелкие не трогаем."""
    if not doc.tables:
        return []
    scored = [(i, len(t.rows)) for i, t in enumerate(doc.tables)]
    best_i, best_n = max(scored, key=lambda x: x[1])
    if best_n >= MIN_ROWS_FOR_MAIN_TABLE:
        return [best_i]
    # fallback: все таблицы с хотя бы 3 строками
    return [i for i, n in scored if n >= 3] or [0]


def apply_layout(doc, layout: dict | None = None, only_main_table: bool = True) -> list[str]:
    """Возвращает предупреждения по документу."""
    standard_cols = (layout or {}).get("grid_cols") or list(DEFAULT_GRID_COLS)
    warnings = diagnose_document(doc, layout)

    indices = _main_table_indices(doc) if only_main_table else list(range(len(doc.tables)))

    for ti in indices:
        table = doc.tables[ti]
        tbl = table._tbl
        grid_cols = _grid_cols_for_table(tbl, standard_cols)
        cumsum = _grid_cumsum(grid_cols)
        table_width = str(cumsum[-1])

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

        jc = tbl_pr.find(qn("w:jc"))
        if jc is None:
            jc = etree.SubElement(tbl_pr, qn("w:jc"))
        jc.set(qn("w:val"), "center")

        old_grid = tbl.find(qn("w:tblGrid"))
        new_grid = _build_tblGrid(grid_cols)
        if old_grid is not None:
            tbl.replace(old_grid, new_grid)
        else:
            tbl_pr.addnext(new_grid)

        for row in table.rows:
            tr = row._tr
            _normalize_row_tr(tr)
            col_idx = 0

            for tc in tr.findall(qn("w:tc")):
                if col_idx >= len(grid_cols):
                    break

                tc_pr = tc.find(qn("w:tcPr"))
                span = 1
                if tc_pr is not None:
                    gs = tc_pr.find(qn("w:gridSpan"))
                    if gs is not None:
                        span = max(1, int(gs.get(qn("w:val"), 1)))
                span = min(span, len(grid_cols) - col_idx)

                cell_w = str(cumsum[col_idx + span] - cumsum[col_idx])
                # Важно: ширину задаём и vMerge-continuation, иначе остаётся старый tcW
                _set_tc_width(tc, cell_w)
                col_idx += span

    return warnings


def main():
    if len(sys.argv) < 2:
        print("Использование: python3 apply_template_layout.py <document.docx> [эталон.docx]")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    template_path = Path(sys.argv[2]) if len(sys.argv) > 2 else resolve_layout_template()

    layout = read_template_layout(template_path)
    doc = Document(os.fspath(input_path))
    print(f"Эталон: {template_path.name}, колонки: {layout['grid_cols']}")
    warns = apply_layout(doc, layout)
    for w in warns:
        print("WARN:", w)

    output_path = input_path.parent / f"{input_path.stem}_layout.docx"
    doc.save(os.fspath(output_path))
    print(f"Сохранён: {output_path}")


if __name__ == "__main__":
    main()
