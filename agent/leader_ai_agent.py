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
    """AI анализирует с учетом положения — шапка vs подвал."""
    
    # Находим макс строку для определения подвала
    max_row = max((c['row'] for c in cells_list), default=0) if cells_list else 0
    mid_row = max_row // 2  # примерная середина
    
    if to_leader == "aniskov":
        target_fio = "Аниськов Владимир Иванович"
        target_title = "Руководитель"  # шапка — короткий
        target_role = "Руководитель проекта СК"  # подвал — полный
        old_hint = "Манджиев/Маджиев, И.о./И.О. Руководителя"
    else:
        target_fio = "Манджиев Игорь Александрович"
        target_title = "И.О. Руководителя"  # шапка — короткий
        target_role = "И.О. Руководителя проекта СК"  # подвал — полный  
        old_hint = "Аниськов, Руководитель/Руководитель проекта СК"

    # Формируем текст для AI с разметкой шапка/подвал
    cells_marked = []
    for c in cells_list:
        section = "ШАПКА" if c['row'] <= mid_row else "ПОДВАЛ"
        cells_marked.append(f"[{section} T{c['table']}R{c['row']}C{c['cell']}]: {c['text']}")

    prompt = f"""Проанализируй ячейки отчёта строительного контроля.

РАЗДЕЛЕНИЕ:
- ШАПКА (верх документа, строки 0-{mid_row}) → короткий заголовок
- ПОДВАЛ (низ документа, строки {mid_row}+) → полная должность

Ищем: {old_hint}

ЦЕЛИ:
- В ШАПКЕ используй: "{target_title}"
- В ПОДВАЛЕ используй: "{target_role}"
- ФИО везде: "{target_fio}"

Ячейки:
{chr(10).join(cells_marked)}

Ответь JSON с replacements по координатам:
{{
  "replacements": [
    {{"old": "точный текст из ячейки", "new": "{target_title}"}},
    {{"old": "точный текст из ячейки", "new": "{target_role}"}},
    {{"old": "точный текст из ячейки", "new": "{target_fio}"}}
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
    """AI говорит что менять, мы меняем в cell.text целиком."""
    try:
        doc = Document(filepath)

        if not doc.tables:
            return False, "Нет таблиц в документе"

        # Собираем все ячейки с текстом и координатами
        cells_data = []
        for t_idx, table in enumerate(doc.tables):
            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                    text = cell.text.strip()
                    if text:
                        cells_data.append({
                            'table': t_idx, 'row': r_idx, 'cell': c_idx,
                            'text': text, 'cell_obj': cell
                        })

        # Отправляем AI список
        ai_result = analyze_with_ai_cells(cells_data, leader)
        
        if "error" in ai_result:
            return False, f"AI ошибка: {ai_result['error']}"

        patterns = ai_result.get("replacements", [])
        
        if not patterns:
            return False, "AI не нашёл паттернов для замены"

        changes = 0
        
        # Ищем ПАТТЕРНЫ в cell.text и заменяем целиком cell.text
        for pattern in patterns:
            old_text = pattern.get('old', '')
            new_text = pattern.get('new', '')
            
            if not old_text or not new_text:
                continue
            
            for cell_info in cells_data:
                if old_text in cell_info['text']:
                    # Заменяем весь текст ячейки
                    cell_info['cell_obj'].text = cell_info['text'].replace(old_text, new_text)
                    changes += 1
                    # Обновляем текст для возможных следующих замен
                    cell_info['text'] = cell_info['cell_obj'].text

        if changes == 0:
            return False, f"AI дал {len(patterns)} паттернов, но не применилось"

        doc.save(filepath)
        return True, f"→ {Path(filepath).name}: замен {changes}"

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
