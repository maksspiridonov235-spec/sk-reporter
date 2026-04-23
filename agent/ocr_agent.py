"""
Агент на базе Ollama + qwen3.5:cloud для анализа и сборки отчётов СК.
Определяет компанию по содержимому документа, вставляет в болванку без ZIP.
"""

import ollama
import re
import json
from copy import deepcopy
from pathlib import Path
from typing import Optional

from docx import Document

KNOWN_COMPANIES = [
    "Евракор", "Лесные технологии", "ЮНС", "НГСК", "Сибитек", "ЭСМ",
    "НГП", "РОСЭКСПО", "ТПС", "ТЭКПРО", "ЮПИ", "УГГ", "ЮГС", "ТВС",
    "НСС", "ОТ и ТБ", "Стройфинансгрупп"
]

MODEL = "qwen3.5:cloud"

SYSTEM_PROMPT = f"""Ты определяешь компанию-подрядчика по тексту отчёта строительного контроля.

Список компаний: {json.dumps(KNOWN_COMPANIES, ensure_ascii=False)}

Правила:
1. Ищи название компании в шапке, подписях, таблицах документа.
2. Верни ТОЛЬКО название из списка выше — без пояснений, без кавычек.
3. Если не нашёл — верни UNKNOWN.
"""


def extract_text(filepath: str) -> str:
    try:
        doc = Document(filepath)
        parts = []
        for para in doc.paragraphs[:50]:
            t = para.text.strip()
            if t:
                parts.append(t)
        for table in doc.tables[:3]:
            for row in table.rows[:5]:
                for cell in row.cells:
                    t = cell.text.strip()
                    if t and len(t) > 2:
                        parts.append(t)
        return "\n".join(parts)
    except Exception as e:
        print(f"[ERROR] extract_text {filepath}: {e}")
        return ""


def detect_company(filepath: str) -> Optional[str]:
    text = extract_text(filepath)
    if not text:
        return None

    try:
        response = ollama.chat(
            model=MODEL,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": f"Определи компанию:\n\n{text[:2000]}"},
            ],
            options={"temperature": 0.0, "num_predict": 30},
        )
        answer = response["message"]["content"].strip()
        # Убираем <think>...</think> если модель добавила
        answer = re.sub(r"<think>.*?</think>", "", answer, flags=re.DOTALL).strip()
        clean = re.sub(r'["\'\.\n]', "", answer).strip()

        for company in KNOWN_COMPANIES:
            if company.lower() in clean.lower() or clean.lower() in company.lower():
                print(f"[AI] {Path(filepath).name} → {company}")
                return company

        print(f"[AI] Не распознана: '{answer}' для {Path(filepath).name}")
        return None

    except Exception as e:
        print(f"[ERROR] ollama: {e}")
        return None


def merge_report_into_template(template_path: str, report_path: str, output_path: str) -> bool:
    """
    Вставляет содержимое report_path в конец template_path.
    Копирует элементы напрямую через python-docx без ZIP-манипуляций.
    Картинки переносятся через document part relationships.
    """
    try:
        master = Document(template_path)
        src = Document(report_path)

        # Копируем relationships (картинки) из src в master
        rid_map: dict[str, str] = {}
        for rel in src.part.rels.values():
            if "image" in rel.reltype:
                img_part = rel.target_part
                new_rid = master.part.relate_to(img_part, rel.reltype)
                rid_map[rel.rId] = new_rid

        # Копируем тело документа
        for element in src.element.body:
            tag = element.tag.split("}")[-1] if "}" in element.tag else element.tag
            if tag == "sectPr":
                continue
            el_copy = deepcopy(element)
            # Патчим rId для картинок
            if rid_map:
                _patch_rids(el_copy, rid_map)
            master.element.body.append(el_copy)

        master.save(output_path)
        return True

    except Exception as e:
        print(f"[ERROR] merge {report_path}: {e}")
        return False


def _patch_rids(element, rid_map: dict) -> None:
    ATTRS = (
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed",
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link",
    )
    for attr in ATTRS:
        val = element.get(attr)
        if val and val in rid_map:
            element.set(attr, rid_map[val])
    for child in element:
        _patch_rids(child, rid_map)
