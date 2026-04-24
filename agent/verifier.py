"""
Агент 3: Верификатор.
Проверяет качество и полноту нормализованного JSON.
"""

import os
import re
import json
from typing import Optional
import anthropic

MODEL = "claude-sonnet-4-6"

PROMPT = """Ты — строгий технический инспектор строительного контроля.
Проверяешь качество заполнения ежедневного отчёта.

Обязательные поля: date, company, object, inspector, work_done.
Желательные поля: weather, conclusion, work_volume.

Верни JSON строго в таком формате:
{
  "ok": true/false,
  "score": число от 0 до 100,
  "missing": ["список обязательных полей которых нет"],
  "warnings": ["список замечаний по качеству заполнения"],
  "summary": "одна фраза — общий вывод о качестве отчёта"
}

Правила:
- ok=true только если все обязательные поля заполнены и нет грубых ошибок
- score: 100 = идеально, 0 = пустой отчёт
- warnings: конкретные замечания (слишком коротко, дата не совпадает с названием файла, и т.п.)
- Верни ТОЛЬКО JSON без пояснений и markdown
"""


def verify(normalized: dict, api_key: Optional[str] = None) -> dict:
    key = api_key or os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        print("[VERIFY] ANTHROPIC_API_KEY не задан — пропускаем верификацию")
        return {"ok": True, "score": 0, "missing": [], "warnings": ["Верификатор отключён"], "summary": ""}

    try:
        client = anthropic.Anthropic(api_key=key)
        response = client.messages.create(
            model=MODEL,
            max_tokens=512,
            system=PROMPT,
            messages=[{
                "role": "user",
                "content": f"Проверь отчёт:\n{json.dumps(normalized, ensure_ascii=False, indent=2)}"
            }],
        )
        raw = response.content[0].text.strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        result = json.loads(raw)
        result["_source_file"] = normalized.get("_source_file")
        score = result.get("score", "?")
        ok = result.get("ok", False)
        print(f"[VERIFY] {'OK' if ok else 'FAIL'} score={score}: {normalized.get('_source_file')}")
        return result
    except Exception as e:
        print(f"[VERIFY] ERROR: {e}")
        return {"ok": False, "score": 0, "missing": [], "warnings": [str(e)], "summary": "Ошибка верификатора"}
