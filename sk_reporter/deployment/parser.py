"""Парсинг суточных отчётов .docx → строки summary."""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

from docx import Document

from sk_reporter.deployment.lookup import normalize_contractor


def normalize_date(text: str) -> str:
    match = re.search(r"(\d{1,2})\.(\d{1,2})\.(\d{2,4})", text)
    if match:
        day, month, year = match.groups()
        year = int(year)
        if year < 100:
            year += 2000
        try:
            return datetime(year, int(month), int(day)).strftime("%d.%m.%Y")
        except ValueError:
            pass
    return text


def extract_from_docx(file_path: str | Path) -> dict[str, str]:
    try:
        doc = Document(str(file_path))
        data = {"Дата": "", "Объект": "", "Инженер СК": "", "Генподрядчик": ""}
        genpodr = ""
        subpodr = ""

        for table in doc.tables:
            for row in table.rows:
                cells = row.cells
                row_text_upper = " ".join(c.text.strip() for c in cells).upper()

                if not data["Дата"]:
                    for i in range(len(cells)):
                        if cells[i].text.strip() == "Дата":
                            for j in range(i + 1, len(cells)):
                                next_text = cells[j].text.strip()
                                if next_text and next_text != "Дата":
                                    data["Дата"] = normalize_date(next_text)
                                    break
                            break

                if not data["Объект"] and "ОБЪЕКТ" in row_text_upper and "СТРАНИЦА" not in row_text_upper:
                    last_obj_idx = -1
                    for i, c in enumerate(cells):
                        if c.text.strip().upper() == "ОБЪЕКТ":
                            last_obj_idx = i
                    if last_obj_idx >= 0:
                        for j in range(last_obj_idx + 1, len(cells)):
                            cell_text = cells[j].text.strip().replace("\n", " ")
                            if cell_text:
                                start = cell_text.find("«")
                                end = cell_text.rfind("»")
                                if start != -1 and end > start:
                                    data["Объект"] = cell_text[start + 1 : end].strip()
                                else:
                                    data["Объект"] = cell_text
                                break

                if not genpodr and "ГЕНПОДРЯДЧИК" in row_text_upper:
                    lt = {c.text.strip() for c in cells if "ГЕНПОДРЯДЧИК" in c.text.upper()}
                    for cell in cells:
                        ct = cell.text.strip()
                        if ct and ct not in lt:
                            genpodr = ct
                            break

                if not subpodr and any(c.text.strip().upper() == "СУБПОДРЯДЧИК" for c in cells):
                    lt = {c.text.strip() for c in cells if c.text.strip().upper() == "СУБПОДРЯДЧИК"}
                    for cell in cells:
                        ct = cell.text.strip()
                        if ct and ct not in lt:
                            subpodr = ct
                            break

            if not data["Инженер СК"] and len(table.rows) >= 2:
                val = table.rows[-2].cells[0].text.strip()
                if val:
                    data["Инженер СК"] = val

        raw = genpodr if genpodr else subpodr
        data["Генподрядчик"] = normalize_contractor(raw)
        return data

    except Exception as exc:
        return {
            "Дата": f"Ошибка: {exc}",
            "Объект": "",
            "Инженер СК": "",
            "Генподрядчик": "",
        }
