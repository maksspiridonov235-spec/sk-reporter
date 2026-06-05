"""
Агент инъекции: берёт оригинальный docx + исправленный текст от check_agent
и вставляет его в ячейки-заголовки секции СК:
«Описание действий», «Участок, ПК», «Ссылка».
"""

import shutil
import tempfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from sk_reporter.agent.report_parts import parse_parts
from sk_reporter.agent.sk_extract import find_sk_section_header_cells

FIXED_DOWNLOAD_SUFFIX = "_исправлен.docx"


def _write_lines_to_cell(cell, lines: list):
    """Заменяет содержимое ячейки, сохраняя первый параграф-заголовок."""
    if not lines:
        return
    tc = cell._tc
    paras = tc.findall(qn("w:p"))
    for p in paras[1:]:
        tc.remove(p)
    for line in lines:
        if line.strip():
            cell.add_paragraph(line)


def inject_into_docx(filepath: str, corrected_text: str, source_filename: str) -> dict:
    stem = Path(source_filename).stem
    try:
        print(f"[INJECT_AGENT] === FULL corrected_text ===\n{corrected_text[:500]}\n... (truncated) ===")
        part1_lines, part2_lines, part3_lines, part4_lines = parse_parts(corrected_text)
        print(
            f"[INJECT_AGENT] parsed part1={len(part1_lines)} part2={len(part2_lines)} "
            f"part3={len(part3_lines)} part4={len(part4_lines)} lines"
        )

        if not part1_lines and not part2_lines:
            return {
                "ok": False,
                "error": "Не удалось распарсить ЧАСТЬ 1 / ЧАСТЬ 2 из ответа агента",
                "docx_path": None,
            }

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir) / Path(filepath).name
            shutil.copy2(filepath, tmp_path)

            doc = Document(str(tmp_path))
            header_cells, coords = find_sk_section_header_cells(doc)
            if not header_cells or "description" not in header_cells:
                return {
                    "ok": False,
                    "error": "Не найдена строка заголовков секции СК в документе",
                    "docx_path": None,
                }

            print(f"[INJECT_AGENT] Found SK header row at {coords}")
            desc_lines = list(part1_lines)
            if part2_lines:
                if desc_lines:
                    desc_lines.append("")
                desc_lines.extend(part2_lines)
            _write_lines_to_cell(header_cells["description"], desc_lines)
            print("[INJECT_AGENT] Wrote parts 1+2 to 'Описание действий' cell")

            if part3_lines and "location" in header_cells:
                _write_lines_to_cell(header_cells["location"], part3_lines)
                print("[INJECT_AGENT] Wrote part 3 to 'Участок, ПК' cell")
            if part4_lines and "reference" in header_cells:
                _write_lines_to_cell(header_cells["reference"], part4_lines)
                print("[INJECT_AGENT] Wrote part 4 to 'Ссылка' cell")

            dest = Path(filepath).resolve()
            doc.save(str(dest))

        print(f"[INJECT_AGENT] Saved to upload: {dest}")
        return {
            "ok": True,
            "docx_path": str(dest),
            "download_name": f"{stem}{FIXED_DOWNLOAD_SUFFIX}",
        }

    except Exception as e:
        import traceback

        print(f"[INJECT_AGENT] error: {e}\n{traceback.format_exc()}")
        return {"ok": False, "error": str(e), "docx_path": None}
