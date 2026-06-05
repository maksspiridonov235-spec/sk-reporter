"""Извлечение секции СК и работ из docx."""

from docx import Document

from sk_reporter.agent.report_parts import (
    parse_numbered_items,
    parse_works_from_part1_lines,
    split_part1_part2_text,
)


def _cell_header_label(cell) -> str:
    text = cell.text.strip()
    return text.split("\n")[0].strip() if text else ""


def _cell_body_text(cell) -> str:
    text = cell.text.strip()
    lines = text.split("\n")
    if lines and _cell_header_label(cell) == lines[0].strip():
        return "\n".join(lines[1:]).strip()
    return text


def find_sk_section_header_cells(doc: Document):
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


def extract_sk_section(filepath: str) -> str:
    """Строки секции СК с колонками Описание / Участок, ПК / Ссылка."""
    try:
        doc = Document(filepath)
        lines = []
        for table in doc.tables:
            in_section = False
            for ri, row in enumerate(table.rows):
                c0 = row.cells[0].text.strip()
                if c0 == "Описание действий" and ri > 0:
                    prev = table.rows[ri - 1].cells[0].text.strip().upper()
                    if "СТРОИТЕЛЬНОГО КОНТРОЛЯ" in prev:
                        in_section = True
                        lines.append("--- Секция строительного контроля (таблица) ---")
                if not in_section:
                    continue
                if ri > 0 and c0.upper().startswith("РЕЗУЛЬТАТ ДУБЛИРУЮЩЕГО"):
                    break
                seen = set()
                cols = []
                for cell in row.cells:
                    tc_id = id(cell._tc)
                    if tc_id in seen:
                        continue
                    seen.add(tc_id)
                    cols.append(cell.text.strip().replace("\n", " | "))
                if any(c.strip() for c in cols):
                    lines.append(" | ".join(cols))
        return "\n".join(lines)
    except Exception as e:
        print(f"[SK_EXTRACT] extract_sk_section error: {e}")
        return ""


def extract_sk_original(filepath: str) -> dict:
    """
    Работы и тексты из ячеек-заголовков секции СК.
    works[]: num, title, volumes, description, location, reference, zero_daily
    """
    try:
        doc = Document(filepath)
        header_cells, _ = find_sk_section_header_cells(doc)
        if not header_cells:
            return {"ok": False, "error": "Секция СК не найдена", "works": []}

        desc_body = _cell_body_text(header_cells["description"])
        part1_text, part2_text = split_part1_part2_text(desc_body)
        part1_lines = [l for l in part1_text.splitlines() if l.strip()]
        part2_lines = [l for l in part2_text.splitlines() if l.strip()]

        works = parse_works_from_part1_lines(part1_lines)
        descriptions = parse_numbered_items(part2_lines)

        locations: dict[int, str] = {}
        references: dict[int, str] = {}
        if "location" in header_cells:
            loc_lines = [l for l in _cell_body_text(header_cells["location"]).splitlines() if l.strip()]
            locations = parse_numbered_items(loc_lines)
        if "reference" in header_cells:
            ref_lines = [l for l in _cell_body_text(header_cells["reference"]).splitlines() if l.strip()]
            references = parse_numbered_items(ref_lines)

        for work in works:
            n = work["num"]
            work["description"] = descriptions.get(n, "")
            work["location"] = locations.get(n, "")
            work["reference"] = references.get(n, "")
            daily = work["volumes"].get("Объем за сутки", "")
            from sk_reporter.agent.report_parts import is_zero_daily_volume

            work["zero_daily"] = is_zero_daily_volume(daily)

        return {
            "ok": True,
            "works": works,
            "active_works": [w for w in works if not w.get("zero_daily")],
            "description_body": desc_body,
            "sk_section": extract_sk_section(filepath),
        }
    except Exception as e:
        print(f"[SK_EXTRACT] extract_sk_original error: {e}")
        return {"ok": False, "error": str(e), "works": [], "active_works": []}
