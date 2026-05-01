"""
Агент инъекции: берёт оригинальный docx + исправленный текст от check_agent,
парсит ЧАСТЬ 1 и ЧАСТЬ 2 и вставляет их в нужные ячейки таблицы через python-docx.
"""

import re
import shutil
import tempfile
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def _clear_cell(cell):
    for para in cell.paragraphs[1:]:
        p = para._element
        p.getparent().remove(p)
    cell.paragraphs[0].clear()


def _add_paragraph(cell, text: str, first: bool = False):
    if first:
        para = cell.paragraphs[0]
        para.clear()
    else:
        para = cell.add_paragraph()
    para.add_run(text)
    return para


def _find_target_cells(doc: Document):
    part1_cell = None
    part2_cell = None
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                txt = cell.text
                if "Инспекционный контроль" in txt and part1_cell is None:
                    part1_cell = cell
                if ("Наряд-допуск проверен" in txt or "Наряд - допуск проверен" in txt) and part2_cell is None:
                    part2_cell = cell
    return part1_cell, part2_cell


def _parse_corrected(corrected_text: str):
    part1_lines = []
    part2_lines = []

    part1_match = re.search(
        r"ЧАСТЬ\s*1[^\n]*\n(.*?)(?=ЧАСТЬ\s*2|$)", corrected_text, re.DOTALL | re.IGNORECASE
    )
    part2_match = re.search(
        r"ЧАСТЬ\s*2[^\n]*\n(.*?)$", corrected_text, re.DOTALL | re.IGNORECASE
    )

    if part1_match:
        part1_lines = [l.rstrip() for l in part1_match.group(1).strip().splitlines()]
    if part2_match:
        part2_lines = [l.rstrip() for l in part2_match.group(1).strip().splitlines()]

    return part1_lines, part2_lines


def _write_lines_to_cell(cell, lines: list):
    _clear_cell(cell)
    first = True
    for line in lines:
        _add_paragraph(cell, line, first=first)
        first = False


def inject_corrections(original_html: str, corrected_text: str, source_filename: str) -> dict:
    filepath = _find_source_docx(source_filename)
    if not filepath:
        return {"ok": False, "error": f"Исходный docx не найден: {source_filename}", "docx_path": None}

    return inject_into_docx(filepath, corrected_text, source_filename)


def _find_source_docx(source_filename: str) -> str | None:
    import tempfile
    work_dir = Path(tempfile.gettempdir()) / "sk_reports_work" / "uploads"
    candidate = work_dir / source_filename
    if candidate.exists():
        return str(candidate)
    return None


def inject_into_docx(filepath: str, corrected_text: str, source_filename: str) -> dict:
    stem = Path(source_filename).stem
    try:
        part1_lines, part2_lines = _parse_corrected(corrected_text)

        if not part1_lines and not part2_lines:
            return {"ok": False, "error": "Не удалось распарсить ЧАСТЬ 1 / ЧАСТЬ 2 из ответа агента", "docx_path": None}

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir) / Path(filepath).name
            shutil.copy2(filepath, tmp_path)

            doc = Document(str(tmp_path))
            part1_cell, part2_cell = _find_target_cells(doc)

            if part1_cell and part1_lines:
                _write_lines_to_cell(part1_cell, part1_lines)

            if part2_cell and part2_lines:
                _write_lines_to_cell(part2_cell, part2_lines)

            output_dir = Path(__file__).parent.parent / "output"
            output_dir.mkdir(exist_ok=True)
            final_path = output_dir / f"{stem}_исправлен.docx"
            doc.save(str(final_path))

        print(f"[INJECT_AGENT] Saved: {final_path}")
        return {"ok": True, "docx_path": str(final_path)}

    except Exception as e:
        print(f"[INJECT_AGENT] error: {e}")
        return {"ok": False, "error": str(e), "docx_path": None}

