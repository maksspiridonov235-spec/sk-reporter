"""Парсинг титульного листа (ТЛ) проекта из docx — таблицы как в Word."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from docx import Document


def _row_cells(row) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for cell in row.cells:
        text = cell.text.strip().replace("\n", " ")
        if not text or text in seen:
            continue
        seen.add(text)
        out.append(text)
    return out


def parse_tl_docx(path: Path) -> dict[str, Any]:
    doc = Document(str(path))
    tables: list[dict[str, Any]] = []
    for table in doc.tables:
        rows: list[list[str]] = []
        for row in table.rows:
            cells = _row_cells(row)
            if cells:
                rows.append(cells)
        if rows:
            tables.append({"rows": rows})
    return {"source": path.name, "tables": tables}
