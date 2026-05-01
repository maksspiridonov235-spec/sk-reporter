"""
AI Агент переключения руководителя через Ollama.
"""

import re
from docx import Document
from pathlib import Path
from typing import Literal, List, Tuple


def switch_leader_ai(filepaths: List[str], leader: Literal["aniskov", "mandzhiev"]) -> Tuple[bool, str]:
    """Обрабатывает ВСЕ загруженные файлы."""
    
    # Определяем направление замены
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
    
    results = []
    success_count = 0
    total_changes = 0
    
    for filepath in filepaths:
        try:
            doc = Document(filepath)
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
            
            if changes > 0:
                doc.save(filepath)
                success_count += 1
                total_changes += changes
                results.append(f"-> {Path(filepath).name}: замен {changes}")
            else:
                results.append(f"-> {Path(filepath).name}: нет изменений")
                
        except Exception as e:
            results.append(f"-> {Path(filepath).name}: ошибка - {e}")
    
    # Формируем вывод: сначала детали, потом итог
    if success_count > 0:
        output = "\n".join(results) + f"\nОбработано: {success_count}/{len(filepaths)} файлов, замен: {total_changes}"
        return True, output
    else:
        return False, "Ни один файл не обработан"


def detect_current_leader(filepath: str) -> str:
    try:
        doc = Document(filepath)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text = cell.text.strip().lower()
                    if "манджиев" in text:
                        return "mandzhiev"
                    elif "аниськов" in text:
                        return "aniskov"
        return "unknown"
    except:
        return "unknown"
