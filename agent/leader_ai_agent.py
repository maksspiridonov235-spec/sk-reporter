"""AI Агент переключения руководителя через Ollama.
Использует LLM для анализа и замены.
"""

import ollama
from docx import Document
from pathlib import Path
from typing import Literal

MODEL = "gemma4:31b-cloud"

SYSTEM_PROMPT = """Ты - агент для редактирования отчётов строительного контроля.

Твоя задача: найти и заменить данные руководителя в документе.

ПРАВИЛА:
1. Найди все упоминания текущего руководителя (ФИО, должность)
2. Замени на нового руководителя
3. Сохрани форматирование документа

СТАРЫЙ РУКОВОДИТЕЛЬ (заменить):
- ФИО: Аниськов Владимир Иванович
- Должность: Руководитель проекта СК
- Заголовок: Руководитель

НОВЫЙ РУКОВОДИТЕЛЬ (вставить):
- ФИО: Манджиев Игорь Александрович
- Должность: И.О. Руководителя проекта СК
- Заголовок: И.О. Руководителя

Ответь JSON:
{
  "found": ["список найденных текстов для замены"],
  "confidence": 0-100
}"""


def extract_text_from_docx(filepath: str) -> str:
    """Извлекает текст из DOCX для анализа."""
    try:
        doc = Document(filepath)
        parts = []

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text = cell.text.strip()
                    if text and len(text) > 2:
                        parts.append(text)

        return "\n".join(parts[:100])
    except Exception as e:
        return f"Error: {e}"


def analyze_with_ai(filepath: str, to_leader: str) -> dict:
    """Анализирует документ через Ollama."""

    if to_leader == "aniskov":
        old_fio = "Манджиев Игорь Александрович"
        new_fio = "Аниськов Владимир Иванович"
        old_title = "И.О. Руководителя"
        new_title = "Руководитель"
        old_project = "И.О. Руководителя проекта СК"
        new_project = "Руководитель проекта СК"
    else:
        old_fio = "Аниськов Владимир Иванович"
        new_fio = "Манджиев Игорь Александрович"
        old_title = "Руководитель"
        new_title = "И.О. Руководителя"
        old_project = "Руководитель проекта СК"
        new_project = "И.О. Руководителя проекта СК"

    prompt = f"""Проанализируй текст и найди все места для замены руководителя.

ЗАМЕНИТЬ:
- "{old_fio}" → "{new_fio}"
- "{old_title}" → "{new_title}"
- "{old_project}" → "{new_project}"

Текст:
{extract_text_from_docx(filepath)}

Ответь JSON:
{{
  "found_patterns": ["..."],
  "confidence": 95
}}"""

    try:
        response = ollama.chat(
            model=MODEL,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": prompt}
            ],
            options={"temperature": 0.0, "num_predict": 500},
        )

        answer = response["message"]["content"]

        import json
        import re

        json_match = re.search(r'\{.*\}', answer, re.DOTALL)
        if json_match:
            return json.loads(json_match.group())

        return {"error": "No JSON found", "raw": answer}

    except Exception as e:
        return {"error": str(e)}


def _switch_single_file(filepath: str, leader: Literal["aniskov", "mandzhiev"]) -> tuple[bool, str]:
    """Обрабатывает один файл."""
    try:
        doc = Document(filepath)

        if not doc.tables:
            return False, "Нет таблиц в документе"

        if leader == "aniskov":
            old_fio = "Манджиев Игорь Александрович"
            new_fio = "Аниськов Владимир Иванович"
            old_title = "И.О. Руководителя"
            new_title = "Руководитель"
            old_project = "И.О. Руководителя проекта СК"
            new_project = "Руководитель проекта СК"
        else:
            old_fio = "Аниськов Владимир Иванович"
            new_fio = "Манджиев Игорь Александрович"
            old_title = "Руководитель"
            new_title = "И.О. Руководителя"
            old_project = "Руководитель проекта СК"
            new_project = "И.О. Руководителя проекта СК"

        changes = 0

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    original = cell.text.strip()
                    new_text = original

                    if old_fio in original:
                        new_text = new_text.replace(old_fio, new_fio)

                    if old_project in original:
                        new_text = new_text.replace(old_project, new_project)

                    if original == old_title:
                        new_text = new_title

                    if new_text != original:
                        cell.text = new_text
                        changes += 1

        if changes == 0:
            ai_result = analyze_with_ai(filepath, leader)
            return False, f"Прямая замена не сработала. AI: {ai_result}"

        doc.save(filepath)

        filename = Path(filepath).name
        return True, f"→ {filename}: замен {changes}"

    except Exception as e:
        filename = Path(filepath).name
        return False, f"→ {filename}: ошибка - {str(e)}"


def switch_leader(filepath: str, leader: Literal["aniskov", "mandzhiev"]) -> tuple[bool, str]:
    """Переключает руководителя в одном файле (совместимость со старым API)."""
    return _switch_single_file(filepath, leader)


def switch_leader_ai(filepaths: list, leader: Literal["aniskov", "mandzhiev"]) -> tuple[bool, str]:
    """Обрабатывает список файлов."""
    if not filepaths:
        return False, "Нет файлов для обработки"
    
    results = []
    success_count = 0
    total_changes = 0
    
    for filepath in filepaths:
        ok, msg = _switch_single_file(filepath, leader)
        results.append(msg)
        if ok:
            success_count += 1
            try:
                if "замен " in msg:
                    changes_str = msg.split("замен ")[-1].strip()
                    total_changes += int(changes_str)
            except:
                pass
    
    if success_count == 0:
        return False, "Ни один файл не обработан: " + "; ".join(results)
    
    output = "\n".join(results)
    output += f"\nОбработано: {success_count}/{len(filepaths)} файлов, замен: {total_changes}"
    return True, output
