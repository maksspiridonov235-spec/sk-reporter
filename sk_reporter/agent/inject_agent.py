"""
Агент инъекции: берёт оригинальный docx + исправленный текст от check_agent
и вставляет его в ячейки-заголовки секции СК:
«Описание действий», «Участок, ПК», «Ссылка».
"""

import re
import shutil
import tempfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

MODEL = "gemma4:31b-cloud"

FIXED_DOWNLOAD_SUFFIX = "_исправлен.docx"


def _split_part(search_text: str, num: int, next_num: int | None) -> list[str]:
    if next_num is not None:
        pattern = rf"ЧАСТЬ\s*{num}[^:\n]*[:\n](.*?)(?=ЧАСТЬ\s*{next_num}\b)"
    else:
        pattern = rf"ЧАСТЬ\s*{num}[^:\n]*[:\n](.*?)$"
    m = re.search(pattern, search_text, re.DOTALL | re.IGNORECASE)
    if not m:
        return []
    return [l.rstrip() for l in m.group(1).strip().splitlines()]


def _parse_parts(corrected_text: str):
    """Parse parts 1–4 from LLM output."""
    cleaned = re.sub(r"\*\*([^*]+)\*\*", r"\1", corrected_text)

    section_match = re.search(
        r"##\s*ИСПРАВЛЕННЫЙ\s*ОТЧЁТ[^\n]*\n(.*?)$", cleaned, re.DOTALL | re.IGNORECASE
    )
    search_text = section_match.group(1).strip() if section_match else cleaned.strip()

    part1_lines = _split_part(search_text, 1, 2)
    part2_lines = _split_part(search_text, 2, 3)
    part3_lines = _split_part(search_text, 3, 4)
    part4_lines = _split_part(search_text, 4, None)

    if not part1_lines and not part2_lines:
        p2_start = re.search(
            r"(Наряд.допуск|Работы ведутся)", search_text, re.IGNORECASE
        )
        if p2_start:
            raw_part1 = search_text[: p2_start.start()].strip()
            raw_part2 = search_text[p2_start.start() :].strip()
            raw_part1 = re.sub(r"^ЧАСТЬ\s*1[^\n]*\n", "", raw_part1, flags=re.IGNORECASE).strip()
            part1_lines = [l.rstrip() for l in raw_part1.splitlines()] if raw_part1 else []
            part2_lines = [l.rstrip() for l in raw_part2.splitlines()] if raw_part2 else []

    print(
        f"[INJECT_AGENT] parsed part1={len(part1_lines)} part2={len(part2_lines)} "
        f"part3={len(part3_lines)} part4={len(part4_lines)} lines"
    )
    return part1_lines, part2_lines, part3_lines, part4_lines


def _cell_header_label(cell) -> str:
    text = cell.text.strip()
    return text.split("\n")[0].strip() if text else ""


def _find_sk_section_header_cells(doc: Document):
    """Ячейки-заголовки строки секции СК: описание, участок, ссылка."""
    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            if row.cells[0].text.strip() != "Описание действий":
                continue
            if ri > 0:
                prev = table.rows[ri - 1].cells[0].text.strip().upper()
                if "СТРОИТЕЛЬНОГО КОНТРОЛЯ" not in prev:
                    continue
            cells = {"description": row.cells[0]}
            seen = {id(row.cells[0]._tc)}
            for cell in row.cells:
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                label = _cell_header_label(cell)
                if label.startswith("Участок"):
                    cells["location"] = cell
                elif label == "Ссылка":
                    cells["reference"] = cell
            return cells, (ti, ri)
    return None, None


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
        part1_lines, part2_lines, part3_lines, part4_lines = _parse_parts(corrected_text)

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
            header_cells, coords = _find_sk_section_header_cells(doc)
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
