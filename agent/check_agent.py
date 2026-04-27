"""
Агент проверки описания действий в отчётах строительного контроля.
Проверяет качество и полноту описания того, что инженер реально проинспектировал.
Не проверяет объемы, только содержимое.
"""

from pathlib import Path
from docx import Document

MODEL = "gemma4:31b-cloud"

SYSTEM_PROMPT = """Ты — ведущий инженер строительного контроля. Проверяешь качество описания действий в отчётах.

КРИТЕРИИ ПРОВЕРКИ описания (раздел "Описание действий"):
1. Указано ЧТО именно проверялось (параметры, размеры, качество, геометрия и т.д.)
2. Указаны СТАНДАРТЫ/ГОСТ/СНИП по которым проверялось
3. Указан РЕЗУЛЬТАТ проверки (принято/не принято/требует переделки)
4. Конкретные ПРИМЕРЫ или ЦИФРЫ (а не общие фразы)

ЗАДАЧА: найти неполные или пустые описания, где не ясно что реально проверяли.

ФОРМАТ ОТВЕТА - для каждого недостатка:

**РАБОТА:** [название]
**ПРОБЛЕМА:** [в чем недостаток]
**НУЖНО ДОБАВИТЬ:** [что должно быть написано]

Если описание полное и конкретное - не пиши о нём.
Не трогай объемы, даты, названия - только содержимое проверки."""


def extract_full_text(filepath: str) -> str:
    """
    Извлекает весь текст из DOCX файла.
    """
    try:
        doc = Document(filepath)
        parts = []

        # Вытаскиваем все параграфы
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                parts.append(text)

        # Вытаскиваем все таблицы
        for table in doc.tables:
            for row in table.rows:
                row_text = " | ".join(cell.text.strip() for cell in row.cells)
                if row_text.strip():
                    parts.append(row_text)

        return "\n".join(parts)
    except Exception as e:
        print(f"[CHECK_AGENT] extract_full_text error: {e}")
        return ""


def check_report(filepath: str) -> dict:
    """
    Проверяет качество описания действий в отчете.
    """
    filename = Path(filepath).name

    # Извлекаем весь текст
    full_text = extract_full_text(filepath)

    if not full_text:
        return {
            "ok": False,
            "report": "Не удалось прочитать содержимое отчета",
            "_source_file": filename,
        }

    # Создаем промпт для LLM
    user_prompt = f"""Проверь описание действий в этом отчете:

---ТЕКСТ ОТЧЕТА---
{full_text}
---КОНЕЦ ТЕКСТА---

ПРОВЕРЯЙ ТОЛЬКО содержимое раздела "Описание действий" (что инженер написал что проинспектировал).

Ищи НЕПОЛНЫЕ описания где:
- Не понятно ЧТО проверялось (только названия работ)
- Не указаны СТАНДАРТЫ/ГОСТ
- Нет РЕЗУЛЬТАТА проверки
- Нет КОНКРЕТНЫХ ПРИМЕРОВ или ЦИФР
- Общие фразы вместо описания

Для каждого недостатка:
**РАБОТА:** [название]
**ПРОБЛЕМА:** [в чем недостаток в описании]
**НУЖНО ДОБАВИТЬ:** [конкретно что написать]

Если описания полные и конкретные - напиши: Описания действий в порядке."""

    # Вызываем ollama
    try:
        import ollama
        response = ollama.chat(
            model=MODEL,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt}
            ],
            stream=False,
        )

        report_text = response.get("message", {}).get("content", "").strip()

        if not report_text:
            return {
                "ok": False,
                "report": "Ошибка: пустой ответ модели",
                "_source_file": filename,
            }

        # Проверяем есть ли проблемы
        has_issues = "в порядке" not in report_text.lower()

        result = {
            "ok": not has_issues,
            "report": report_text,
            "_source_file": filename,
        }

        print(f"[CHECK_AGENT] {'OK' if result['ok'] else 'ISSUES'}: {filename}")
        return result

    except Exception as e:
        print(f"[CHECK_AGENT] Error calling ollama: {e}")
        return {
            "ok": False,
            "report": f"Ошибка проверки: {str(e)}",
            "_source_file": filename,
        }
