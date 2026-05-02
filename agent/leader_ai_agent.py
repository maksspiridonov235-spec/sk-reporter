"""AI Агент переключения руководителя через Ollama.
Использует LLM для анализа и замены.
"""

import ollama
from docx import Document
from pathlib import Path
from typing import Literal

MODEL = "gemma4:31b-cloud"

SYSTEM_PROMPT = """Ты - агент для редактирования отчётов строительного контроля.

Твоя задача: проанализировать документ и найти ВСЕ варианты написания данных руководителя.

ПРАВИЛА:
1. Ищи все варианты написания: с разным регистром (И.О./И.о./и.о.), с опечатками, сокращениями
2. Найди ФИО, должность и заголовок даже если они написаны нестандартно
3. Верни СПИСОК конкретных текстовых паттернов, которые нужно заменить

СТАРЫЙ РУКОВОДИТЕЛЬ (искать все варианты):
- ФИО: Аниськов Владимир Иванович (и опечатки типа Анисков, Аниськов В.И., и т.д.)
- Должность: Руководитель проекта СК (и варианты: Руководителя проекта СК, и.о. руководителя проекта СК)
- Заголовок: Руководитель (и варианты: Руководителя, и.о. руководителя)

НОВЫЙ РУКОВОДИТЕЛЬ (вставить):
- ФИО: Манджиев Игорь Александрович
- Должность: И.О. Руководителя проекта СК
- Заголовок: И.О. Руководителя

Ответь JSON со списком найденных паттернов для замены:
{
  "replacements": [
    {"old": "текст который нашел", "new": "текст на который меняешь"},
    {"old": "другой вариант", "new": "текст на который меняешь"}
  ],
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
    """Анализирует документ через Ollama и возвращает список замен."""
    
    if to_leader == "aniskov":
        direction = "Манджиев/и.о. → Аниськов/руководитель"
        target_fio = "Аниськов Владимир Иванович"
        target_role = "Руководитель проекта СК"
        target_title = "Руководитель"
    else:
        direction = "Аниськов/руководитель → Манджиев/и.о."
        target_fio = "Манджиев Игорь Александрович"
        target_role = "И.О. Руководителя проекта СК"
        target_title = "И.О. Руководителя"

    prompt = f"""Направление замены: {direction}

Найди в документе ВСЕ варианты написания и верни JSON с replacements.

Target values:
- FIO: {target_fio}
- Role: {target_role}
- Title: {target_title}

Text:
{extract_text_from_docx(filepath)}

Ответь JSON:
{{
  "replacements": [
    {{"old": "найденный текст 1", "new": "{target_fio}"}},
    {{"old": "найденный текст 2", "new": "{target_role}"}},
    {{"old": "найденный текст 3", "new": "{target_title}"}}
  ],
  "confidence": 0-100
}}"""

    try:
        response = ollama.chat(
            model=MODEL,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": prompt}
            ],
            options={"temperature": 0.0, "num_predict": 1000},
        )

        answer = response["message"]["content"]

        import json
        import re

        json_match = re.search(r'\{.*\}', answer, re.DOTALL)
        if json_match:
            result = json.loads(json_match.group())
            # Добавляем fallback replacements если AI ничего не нашел
            if not result.get("replacements"):
                if to_leader == "aniskov":
                    result["replacements"] = [
                        {"old": "Манджиев Игорь Александрович", "new": "Аниськов Владимир Иванович"},
                        {"old": "Маджиев Игорь Александрович", "new": "Аниськов Владимир Иванович"},
                        {"old": "И.О. Руководителя проекта СК", "new": "Руководитель проекта СК"},
                        {"old": "И.о. Руководителя проекта СК", "new": "Руководитель проекта СК"},
                        {"old": "И.О. Руководителя", "new": "Руководитель"},
                        {"old": "И.о. Руководителя", "new": "Руководитель"},
                    ]
                else:
                    result["replacements"] = [
                        {"old": "Аниськов Владимир Иванович", "new": "Манджиев Игорь Александрович"},
                        {"old": "Руководитель проекта СК", "new": "И.О. Руководителя проекта СК"},
                        {"old": "Руководитель", "new": "И.О. Руководителя"},
                    ]
            return result

        return {"error": "No JSON found", "raw": answer}

    except Exception as e:
        return {"error": str(e)}


def _replace_in_runs(cell, old_text: str, new_text: str) -> int:
    """Заменяет текст внутри run'ов ячейки, сохраняя форматирование."""
    changes = 0
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            # Проверяем точное совпадение после нормализации пробелов
            normalized_run = " ".join(run.text.split())  # убираем лишние пробелы
            if old_text in normalized_run or old_text in run.text:
                # Заменяем в оригинальном тексте
                run.text = run.text.replace(old_text, new_text)
                changes += 1
    return changes


def _switch_single_file(filepath: str, leader: Literal["aniskov", "mandzhiev"]) -> tuple[bool, str]:
    """Обрабатывает один файл с использованием AI для поиска паттернов."""
    try:
        doc = Document(filepath)

        if not doc.tables:
            return False, "Нет таблиц в документе"

        # Получаем рекомендации от AI
        ai_result = analyze_with_ai(filepath, leader)
        
        if "error" in ai_result:
            return False, f"AI ошибка: {ai_result['error']}"

        replacements = ai_result.get("replacements", [])
        
        if not replacements:
            return False, "AI не нашёл паттернов для замены"

        changes = 0

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for repl in replacements:
                        old = repl.get("old", "")
                        new = repl.get("new", "")
                        if old and new:
                            changes += _replace_in_runs(cell, old, new)

        if changes == 0:
            return False, f"Найдены паттерны, но замена не сработала. AI: {ai_result}"

        doc.save(filepath)

        filename = Path(filepath).name
        return True, f"→ {filename}: замен {changes} (confidence: {ai_result.get('confidence', 'N/A')}%)"

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
