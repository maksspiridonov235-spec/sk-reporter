"""
Агент 2: Нормализатор.
Приводит сырой JSON от парсера к единому формату.
"""

import os
import re
import json
from typing import Optional
import anthropic

MODEL = "claude-sonnet-4-6"

KNOWN_COMPANIES = [
    "Евракор", "Лесные технологии", "ЮНС", "НГСК", "Сибитек", "ЭСМ",
    "НГП", "РОСЭКСПО", "ТПС", "ТЭКПРО", "ЮПИ", "УГГ", "ЮГС", "ТВС",
    "НСС", "ОТ и ТБ", "Стройфинансгрупп",
]

PROMPT = f"""Ты нормализуешь JSON-данные из отчёта строительного контроля.

Список допустимых компаний: {json.dumps(KNOWN_COMPANIES, ensure_ascii=False)}

Правила нормализации:
1. "date" → формат ДД.ММ.ГГГГ, если дата есть но в другом формате — приведи
2. "company" → замени на точное название из списка выше (по смыслу), если не совпадает — оставь как есть
3. "work_done", "violations", "conclusion" → убери лишние пробелы, переносы, дубли фраз
4. "work_volume" → оставь число + единицу измерения, убери лишнее
5. Все null-поля оставь null, не придумывай данные
6. Верни ТОЛЬКО валидный JSON без пояснений и markdown
"""


def normalize(parsed: dict, api_key: Optional[str] = None) -> dict:
    key = api_key or os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        print("[NORM] ANTHROPIC_API_KEY не задан — пропускаем нормализацию")
        return parsed

    try:
        client = anthropic.Anthropic(api_key=key)
        response = client.messages.create(
            model=MODEL,
            max_tokens=1024,
            system=PROMPT,
            messages=[{
                "role": "user",
                "content": f"Нормализуй:\n{json.dumps(parsed, ensure_ascii=False, indent=2)}"
            }],
        )
        raw = response.content[0].text.strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        result = json.loads(raw)
        result["_source_file"] = parsed.get("_source_file")
        print(f"[NORM] OK: {parsed.get('_source_file')}")
        return result
    except Exception as e:
        print(f"[NORM] ERROR: {e}")
        return parsed
