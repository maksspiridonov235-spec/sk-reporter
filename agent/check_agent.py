"""
Агент проверки загруженных отчётов строительного контроля.
Проверяет объемы работ, удаляет работы с 0 суточным объемом,
проверяет соответствие описаний действий типам работ.
Возвращает отчет об ошибках и рекомендациях (без изменения файлов).
"""

import re
from pathlib import Path
from docx import Document

MODEL = "gemma4:31b-cloud"

SYSTEM_PROMPT = """Ты — помощник инженера, который проверяет и переписывает его отчеты правильно.

ПРАВИЛА ПРОВЕРКИ:
1. ОБЪЕМЫ: накопительный объем должен быть ≤ проектному
2. НУЛЕВЫЕ ОБЪЕМЫ: если суточный объем = 0, работу удалить
3. ОПИСАНИЯ: только описание работы БЕЗ цифр объемов

ФОРМАТ ОТВЕТА - для каждой найденной ОШИБКИ:

**БЫЛО (ошибка):**
[точно что было]

**ИСПРАВИТЬ НА:**
[как правильно, с цифрами]

Если ошибок нет - напиши: Ошибок не найдено.
Будь конкретен и точен."""


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
    Основная функция агента.
    Читает весь текст отчета и проверяет его.
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

    # Создаем промпт для LLM с полным текстом отчета
    user_prompt = f"""Проверь этот отчет строительного контроля:

---ТЕКСТ ОТЧЕТА---
{full_text}
---КОНЕЦ ТЕКСТА---

Проверь по ПРАВИЛАМ:
1. Объемы работ (нет ли превышения накопительного над проектным)
2. Работы с нулевым суточным объемом (их нужно удалить)
3. Описания (не содержат ли лишние цифры объемов)

Для каждой ОШИБКИ покажи:
**БЫЛО:** [ошибка из отчета]
**ИСПРАВИТЬ НА:** [правильный вариант]

Если ошибок нет - напиши только: Ошибок не найдено."""

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

        # Проверяем есть ли ошибки
        has_errors = "ошибок не найдено" not in report_text.lower()

        result = {
            "ok": not has_errors,
            "report": report_text,
            "_source_file": filename,
        }

        print(f"[CHECK_AGENT] {'OK' if result['ok'] else 'ERRORS'}: {filename}")
        return result

    except Exception as e:
        print(f"[CHECK_AGENT] Error calling ollama: {e}")
        return {
            "ok": False,
            "report": f"Ошибка проверки: {str(e)}",
            "_source_file": filename,
        }
