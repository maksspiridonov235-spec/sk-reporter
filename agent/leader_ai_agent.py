"""
AI Агент переключения руководителя через Ollama.
Использует LLM для анализа и замены.
"""

import ollama
from docx import Document
from pathlib import Path
from typing import Literal

MODEL = "qwen3.5:cloud"

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
        
        return "\n".join(parts[:100])  # Первые 100 ячеек
    except Exception as e:
        return f"Error: {e}"


def analyze_with_ai(filepath: str, to_leader: str) -> dict:
    """Анализирует документ через Ollama."""
    
    # Определяем направление замены
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
    
    prompt = f"""Проанализируй текст отчёта и найди все места где нужно заменить руководителя.

ЗАМЕНИТЬ:
- "{old_fio}" → "{new_fio}"
- "{old_title}" → "{new_title}"
- "{old_project}" → "{new_project}"

Текст документа:
{extract_text_from_docx(filepath)}

Ответь только JSON:
{{
  "found_patterns": ["список найденных паттернов"],
  "cells_to_modify": [{"row": 0, "col": 0, "old_text": "...", "new_text": "..."}],
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
        
        # Парсим JSON из ответа
        import json
        import re
        
        # Ищем JSON в ответе
        json_match = re.search(r'\{.*\}', answer, re.DOTALL)
        if json_match:
            return json.loads(json_match.group())
        
        return {"error": "No JSON found", "raw": answer}
        
    except Exception as e:
        return {"error": str(e)}


def switch_leader_ai_single(filepath: str, leader: Literal["aniskov", "mandzhiev"]) -> tuple[bool, str]:
    """
    AI-агент переключения руководителя.
    """
    try:
        doc = Document(filepath)
        
        if not doc.tables:
            return False, "Нет таблиц в документе"
        
        # Определяем направление
        if leader == "aniskov":
            old_fio = "Манджиев Игорь Александрович"
            new_fio = "Аниськов Владимир Иванович"
            old_title = "И.О. Руководителя"
            new_title = "Руководитель"
            old_project = "И.О. Руководителя проекта СК"
            new_project = "Руководитель проекта СК"
            target = "Аниськов В.И."
        else:
            old_fio = "Аниськов Владимир Иванович"
            new_fio = "Манджиев Игорь Александрович"
            old_title = "Руководитель"
            new_title = "И.О. Руководителя"
            old_project = "Руководитель проекта СК"
            new_project = "И.О. Руководителя проекта СК"
            target = "Манджиев И.А."
        
        changes = []
        
        # Заменяем напрямую без AI (быстрее и надёжнее)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    original = cell.text.strip()
                    new_text = original
                    
                    # ФИО
                    if old_fio in original:
                        new_text = new_text.replace(old_fio, new_fio)
                    
                    # Заголовок проекта
                    if old_project in original:
                        new_text = new_text.replace(old_project, new_project)
                    
                    # Общий заголовок (только полное совпадение)
                    if original == old_title:
                        new_text = new_title
                    
                    if new_text != original:
                        cell.text = new_text
                        changes.append(f"{original[:30]}... -> {new_text[:30]}...")
        
        if not changes:
            # AI анализ если прямой замены не было
            ai_result = analyze_with_ai(filepath, leader)
            return False, f"Прямая замена не сработала. AI анализ: {ai_result}"
        
        doc.save(filepath)
        
        return True, f"AI-агент: установлен {target}. Замен: {len(changes)}"
        
    except Exception as e:
        return False, f"AI-агент ошибка: {str(e)}"

def switch_leader_ai(filepaths: list, leader: Literal["aniskov", "mandzhiev"]) -> tuple[bool, str]:
    """Обрабатывает ВСЕ загруженные файлы (10, 50, 100 - любое количество)."""
    results = []
    total_changes = 0
    success_count = 0
    
    for filepath in filepaths:
        try:
            ok, msg = switch_leader_ai_single(filepath, leader)
            results.append(f"{Path(filepath).name}: {msg}")
            if ok:
                success_count += 1
                # Извлекаем число замен из сообщения
                import re
                match = re.search(r'Замен: (\d+)', msg)
                if match:
                    total_changes += int(match.group(1))
        except Exception as e:
            results.append(f"{Path(filepath).name}: Ошибка - {e}")
    
    summary = f"Обработано: {success_count}/{len(filepaths)} файлов, замен: {total_changes}"
    
    # Формируем детальный отчет для журнала
    details = []
    for r in results:
        if ": AI-агент:" in r:
            # Извлекаем имя файла и количество замен
            fname = r.split(":")[0]
            match = re.search(r'Замен: (\d+)', r)
            if match:
                details.append(f"→ {fname}: замен {match.group(1)}")
    
    # Если хотя бы один успешно - возвращаем True
    if success_count > 0:
        return True, summary + "
" + "
".join(details[:10])  # Первые 10 файлов
    else:
        return False, "Ни один файл не обработан"
