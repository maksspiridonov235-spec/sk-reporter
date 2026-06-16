"""Дословное извлечение шести пунктов ОТКК из .doc (textutil, без правок текста)."""

from __future__ import annotations

import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any


def _doc_to_txt(path: Path) -> str:
    path = Path(path)
    if path.suffix.lower() == ".docx":
        from docx import Document

        doc = Document(str(path))
        return "\n".join(p.text for p in doc.paragraphs)

    if sys.platform == "darwin":
        textutil = shutil.which("textutil")
        if textutil:
            proc = subprocess.run(
                [textutil, "-convert", "txt", "-stdout", str(path)],
                capture_output=True,
                text=True,
                errors="replace",
            )
            if proc.returncode == 0 and proc.stdout:
                return proc.stdout

    raise RuntimeError(f"Не удалось прочитать {path.name} дословно (нужен macOS textutil или .docx)")


def _after_label(chunk: str, label: str) -> str:
    if not chunk.startswith(label):
        return chunk
    rest = chunk[len(label) :]
    return rest.lstrip("\x07 \t")


def _visible_normative_line(norm_block: str) -> str:
    """Коды СП/ГОСТ как видны в ячейке Word (после «Статус…», без СНиП из подсказок ссылок)."""
    codes = re.findall(r'\)"\s*((?:СП|ГОСТ|ВСН)\s*[\d][\d.\-]*)', norm_block)
    if not codes:
        codes = re.findall(r"(?:СП|ГОСТ|ВСН)\s*[\d][\d.\-]*", norm_block)
    seen: set[str] = set()
    out: list[str] = []
    for raw in codes:
        code = re.sub(r"\s+", " ", raw.strip())
        key = code.casefold()
        if key not in seen:
            seen.add(key)
            out.append(code)
    return ", ".join(out)


def extract_six_rows_from_doc(path: Path) -> dict[str, Any]:
    """Шесть пунктов карты: label/value как в исходном .doc, без очистки HYPERLINK и таблиц."""
    path = Path(path)
    raw = _doc_to_txt(path)

    m_code = re.search(r"ОТКК\s*[-–—]\s*(\d+)", path.stem, re.I)
    card_id = f"otkk-{int(m_code.group(1))}" if m_code else ""

    code_m = re.search(r"(ОТКК\s*[-–—]\s*\d+)", raw)
    code = code_m.group(1).strip() if code_m else ""

    norm_pos = raw.find("Нормативные документы")
    instr_pos = raw.find("Приборы контроля")
    ctrl_pos = raw.find("Контролируемые параметры и необходимая документация")
    sig_pos = raw.find("Разработал:")

    if norm_pos < 0 or instr_pos < 0 or ctrl_pos < 0 or sig_pos < 0:
        raise ValueError(f"Не найдена структура карты в {path.name}")

    head = raw[:norm_pos]
    head_lines = [ln.strip() for ln in head.splitlines() if ln.strip()]

    title = ""
    for ln in head_lines:
        if ln.startswith("ОТКК"):
            continue
        if ln in ("Наименование предприятия", "Область применения"):
            continue
        if "Операционная технологическая" in ln or "карта" in ln.lower():
            title = ln.strip()
            break
    if not title and len(head_lines) > 1:
        title = head_lines[1]

    enterprise = ""
    scope = ""
    for i, ln in enumerate(head_lines):
        if ln == "Наименование предприятия" and i + 1 < len(head_lines):
            enterprise = head_lines[i + 1]
        if ln == "Область применения" and i + 1 < len(head_lines):
            scope = head_lines[i + 1]

    norm_raw = _after_label(raw[norm_pos:instr_pos], "Нормативные документы")
    norm_value = _visible_normative_line(norm_raw)
    instr_value = _after_label(raw[instr_pos:ctrl_pos], "Приборы контроля:")
    ctrl_value = _after_label(
        raw[ctrl_pos:sig_pos],
        "Контролируемые параметры и необходимая документация",
    )

    rows = [
        {"label": code, "value": title},
        {"label": "Наименование предприятия", "value": enterprise},
        {"label": "Область применения", "value": scope},
        {"label": "Нормативные документы", "value": norm_value},
        {"label": "Приборы контроля", "value": instr_value},
        {"label": "Контролируемые параметры и необходимая документация", "value": ctrl_value},
    ]

    return {
        "id": card_id,
        "code": code,
        "title": title,
        "file": path.name,
        "rows": rows,
        "signature": None,
        "plain_text": raw,
    }
