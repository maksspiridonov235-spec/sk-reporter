"""Сборка ежедневного отчёта инженера (docx) из шаблона + заполненных работ."""

from __future__ import annotations

import shutil
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from docx import Document

from sk_reporter.agent.inject_agent import _find_sk_section_header_cells, _write_lines_to_cell
from sk_reporter.engineer.doc_text import control_snippet_from_tk
from sk_reporter.engineer.tk_catalog import resolve_tk_for_work, tk_text_for_id


@dataclass
class ReportEntry:
    name: str
    unit: str
    project_qty: str
    daily_qty: str
    cumulative_qty: str
    location: str = ""
    reference: str = ""
    stage: str = ""
    object_title: str = ""


def _tk_snippet(work_name: str, project_id: str) -> str:
    tk_id = resolve_tk_for_work(work_name, project_id)
    if not tk_id:
        return ""
    try:
        text = tk_text_for_id(tk_id)
        if not text:
            return ""
        return control_snippet_from_tk(text)
    except Exception as exc:
        return f"[ТК {tk_id}: {exc}]"


def _build_part1(entries: list[ReportEntry]) -> list[str]:
    lines = ["Инспекционный контроль по проведённым работам:"]
    for i, e in enumerate(entries, 1):
        lines.append(f"{i}. {e.name}")
        if e.project_qty:
            lines.append(f"Проектный объем – {e.project_qty} {e.unit}".strip())
        if e.daily_qty:
            lines.append(f"Объем за сутки – {e.daily_qty} {e.unit}".strip())
        if e.cumulative_qty:
            lines.append(f"Накопительный объем – {e.cumulative_qty} {e.unit}".strip())
        lines.append("")
    return lines


def _build_part2(entries: list[ReportEntry], project_id: str) -> list[str]:
    lines = [
        "Наряд-допуск проверен, работы ведутся в соответствии с проектной документацией.",
        "",
    ]
    for i, e in enumerate(entries, 1):
        snippet = _tk_snippet(e.name, project_id)
        if snippet:
            lines.append(f"{i}. {snippet}")
        else:
            lines.append(f"{i}. Выполнен инспекционный контроль: {e.name}.")
    return lines


def build_report_docx(
    template_path: Path,
    output_path: Path,
    entries: list[ReportEntry],
    project_id: str,
    report_date: date | None = None,
) -> Path:
    if not entries:
        raise ValueError("Нет работ для отчёта")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, output_path)

    doc = Document(str(output_path))
    header_cells, coords = _find_sk_section_header_cells(doc)
    if not header_cells or "description" not in header_cells:
        raise ValueError("В шаблоне не найдена секция СК (строка «Описание действий»)")

    part1 = _build_part1(entries)
    part2 = _build_part2(entries, project_id)
    desc_lines = part1 + ([""] if part1 and part2 else []) + part2
    _write_lines_to_cell(header_cells["description"], desc_lines)

    part3 = [e.location or "—" for e in entries]
    part4 = [e.reference or "—" for e in entries]
    if "location" in header_cells:
        _write_lines_to_cell(header_cells["location"], part3)
    if "reference" in header_cells:
        _write_lines_to_cell(header_cells["reference"], part4)

    doc.save(str(output_path))
    _ = report_date  # дата в шапке — отдельный шаг (как в /daily)
    print(f"[ENGINEER] report built: {output_path.name} rows={coords}")
    return output_path
