"""
Агент инъекции: берёт оригинальный docx + исправленный текст от check_agent
и вставляет его в ячейку-заголовок «Описание действий».
"""

import re
import shutil
import tempfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from sk_reporter.paths import output_dir

MODEL = "gemma4:31b-cloud"


def _parse_parts(corrected_text: str):
    """Parse part1 and part2 from LLM output."""
    cleaned = re.sub(r"\*\*([^*]+)\*\*", r"\1", corrected_text)

    section_match = re.search(
        r"##\s*ИСПРАВЛЕННЫЙ\s*ОТЧЁТ[^\n]*\n(.*?)$", cleaned, re.DOTALL | re.IGNORECASE
    )
    search_text = section_match.group(1).strip() if section_match else cleaned.strip()

    part1_match = re.search(
        r"ЧАСТЬ\s*1[^:\n]*[:\n](.*?)(?=ЧАСТЬ\s*2\b|$)", search_text, re.DOTALL | re.IGNORECASE
    )
    part2_match = re.search(
        r"ЧАСТЬ\s*2[^:\n]*[:\n](.*?)$", search_text, re.DOTALL | re.IGNORECASE
    )

    if not part1_match or not part2_match:
        p2_start = re.search(
            r"(Наряд.допуск|Работы ведутся)", search_text, re.IGNORECASE
        )
        if p2_start:
            raw_part1 = search_text[: p2_start.start()].strip()
            raw_part2 = search_text[p2_start.start() :].strip()
            raw_part1 = re.sub(r"^ЧАСТЬ\s*1[^\n]*\n", "", raw_part1, flags=re.IGNORECASE).strip()
            part1_lines = [l.rstrip() for l in raw_part1.splitlines()] if raw_part1 else []
            part2_lines = [l.rstrip() for l in raw_part2.splitlines()] if raw_part2 else []
            print(f"[INJECT_AGENT] fallback parse: part1={len(part1_lines)} lines, part2={len(part2_lines)} lines")
            return part1_lines, part2_lines

    part1_lines = []
    part2_lines = []
    if part1_match:
        part1_lines = [l.rstrip() for l in part1_match.group(1).strip().splitlines()]
    if part2_match:
        part2_lines = [l.rstrip() for l in part2_match.group(1).strip().splitlines()]

    print(f"[INJECT_AGENT] parsed part1={len(part1_lines)} lines, part2={len(part2_lines)} lines")
    return part1_lines, part2_lines


def _find_description_header_cell(doc: Document):
    """Ищет ячейку с заголовком «Описание действий» (секция строительного контроля)."""
    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            cell = row.cells[0]
            if cell.text.strip() != "Описание действий":
                continue
            if ri > 0:
                prev = table.rows[ri - 1].cells[0].text.strip().upper()
                if "СТРОИТЕЛЬНОГО КОНТРОЛЯ" not in prev:
                    continue
            return cell, (ti, ri)
    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            for cell in row.cells:
                if cell.text.strip() == "Описание действий":
                    return cell, (ti, ri)
    return None, None


def _write_parts_to_cell(cell, part1_lines: list, part2_lines: list):
    """Заменяет содержимое ячейки, сохраняя первый параграф «Описание действий»."""
    all_lines = []
    if part1_lines:
        all_lines.extend(part1_lines)
    if part2_lines:
        all_lines.extend(part2_lines)
    if not all_lines:
        return

    tc = cell._tc
    paras = tc.findall(qn("w:p"))
    for p in paras[1:]:
        tc.remove(p)

    for line in all_lines:
        cell.add_paragraph(line)


def inject_into_docx(filepath: str, corrected_text: str, source_filename: str) -> dict:
    stem = Path(source_filename).stem
    try:
        print(f"[INJECT_AGENT] === FULL corrected_text ===\n{corrected_text[:500]}\n... (truncated) ===")
        part1_lines, part2_lines = _parse_parts(corrected_text)

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
            target_cell, coords = _find_description_header_cell(doc)
            if target_cell is None:
                return {
                    "ok": False,
                    "error": "Не найдена ячейка «Описание действий» в документе",
                    "docx_path": None,
                }

            print(f"[INJECT_AGENT] Found 'Описание действий' at {coords}")
            _write_parts_to_cell(target_cell, part1_lines, part2_lines)
            print("[INJECT_AGENT] Wrote corrected parts to 'Описание действий' cell")

            out_dir = output_dir()
            final_path = out_dir / f"{stem}_исправлен.docx"
            doc.save(str(final_path))

        print(f"[INJECT_AGENT] Saved: {final_path}")
        return {"ok": True, "docx_path": str(final_path)}

    except Exception as e:
        import traceback

        print(f"[INJECT_AGENT] error: {e}\n{traceback.format_exc()}")
        return {"ok": False, "error": str(e), "docx_path": None}
