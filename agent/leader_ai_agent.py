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
    """Старый метод для совместимости."""
    return analyze_with_ai_cells([], to_leader)


def analyze_with_ai_cells(cells_list: list, to_leader: str) -> dict:
    """AI анализирует список ячеек и говорит что менять."""
    
    if to_leader == "aniskov":
        target_fio = "Аниськов Владимир Иванович"
        target_role = "Руководитель проекта СК"
        target_title = "Руководитель"
    else:
        target_fio = "Манджиев Игорь Александрович"
        target_role = "И.О. Руководителя проекта СК"
        target_title = "И.О. Руководителя"

    # Формируем текст для AI
    cells_text = "\n".join([
        f"[{c['table']},{c['row']},{c['cell']}]: {c['text']}" 
        for c in cells_list
    ])

    prompt = f"""Проанализируй ячейки таблицы и найди те, где нужно заменить руководителя.

Target:
- FIO: {target_fio}
- Role: {target_role}
- Title: {target_title}

Cells:
{cells_text}

Ответь JSON с координатами ячеек и новым текстом:
{{
  "cell_replacements": [
    {{"table": 0, "row": 1, "cell": 2, "new_text": "{target_fio}"}},
    {{"table": 0, "row": 3, "cell": 1, "new_text": "{target_role}"}}
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
            options={"temperature": 0.0, "num_predict": 1500},
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
    """AI делает всё: находит ячейки, говорит что менять, мы меняем целиком."""
    try:
        doc = Document(filepath)

        if not doc.tables:
            return False, "Нет таблиц в документе"

        # Собираем все ячейки с текстом
        cells_with_text = []
        for t_idx, table in enumerate(doc.tables):
            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                    text = cell.text.strip()
                    if text:
                        cells_with_text.append({
                            'table': t_idx,
                            'row': r_idx,
                            'cell': c_idx,
                            'text': text,
                            'cell_obj': cell
                        })

        # Отправляем AI список ячеек
        ai_result = analyze_with_ai_cells(cells_with_text, leader)
        
        if "error" in ai_result:
            return False, f"AI ошибка: {ai_result['error']}"

        cell_replacements = ai_result.get("cell_replacements", [])
        
        if not cell_replacements:
            return False, "AI не нашёл ячеек для замены"

        changes = 0
        
        # Применяем замены по координатам
        for repl in cell_replacements:
            t = repl.get('table')
            r = repl.get('row')
            c = repl.get('cell')
            new_text = repl.get('new_text')
            
            # Находим ячейку и меняем целиком
            for cell_info in cells_with_text:
                if (cell_info['table'] == t and 
                    cell_info['row'] == r and 
                    cell_info['cell'] == c):
                    cell_info['cell_obj'].text = new_text
                    changes += 1
                    break

        if changes == 0:
            return False, f"AI дал {len(cell_replacements)} замен, но не применилось"

        doc.save(filepath)
        return True, f"→ {Path(filepath).name}: замен {changes} ячеек"

    except Exception as e:
        return False, f"→ {Path(filepath).name}: ошибка - {str(e)}"


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
