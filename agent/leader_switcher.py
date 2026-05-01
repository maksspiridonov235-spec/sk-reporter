"""
Умный агент переключения руководителя в отчёте СК.
Ищет и заменяет по содержимому.
"""

from docx import Document
from typing import Literal, List, Tuple


LEADERS = {
    "aniskov": {
        "fio": "Аниськов Владимир Иванович",
        "project_title": "Руководитель проекта СК",
        "title": "Руководитель",
    },
    "mandzhiev": {
        "fio": "Манджиев Игорь Александрович",
        "project_title": "И.О. Руководителя проекта СК",
        "title": "И.О. Руководителя",
    }
}


def switch_leader(filepath: str, leader: Literal["aniskov", "mandzhiev"]) -> Tuple[bool, str]:
    try:
        doc = Document(filepath)
        
        if not doc.tables:
            return False, "В документе нет таблиц"
        
        # Кого ставим
        new = LEADERS[leader]
        target_name = "Аниськов В.И." if leader == "aniskov" else "Манджиев И.А."
        
        changes = []
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text = cell.text.strip()
                    
                    # Ищем Аниськова - меняем на Манджиева (или наоборот)
                    if "Аниськов Владимир Иванович" in text:
                        cell.text = text.replace("Аниськов Владимир Иванович", new["fio"])
                        changes.append("ФИО")
                    
                    if "Манджиев Игорь Александрович" in text:
                        cell.text = text.replace("Манджиев Игорь Александрович", new["fio"])
                        changes.append("ФИО")
                    
                    # Заголовок проекта
                    if "Руководитель проекта СК" in text and "И.О." not in text:
                        cell.text = text.replace("Руководитель проекта СК", new["project_title"])
                        changes.append("Проект")
                    
                    if "И.О. Руководителя проекта СК" in text:
                        cell.text = text.replace("И.О. Руководителя проекта СК", new["project_title"])
                        changes.append("Проект")
                    
                    # Общий заголовок в договоре
                    if text == "Руководитель":
                        cell.text = new["title"]
                        changes.append("Заголовок")
                    
                    if text == "И.О. Руководителя":
                        cell.text = new["title"]
                        changes.append("Заголовок")
        
        if not changes:
            return False, "Не найдено полей для замены"
        
        doc.save(filepath)
        return True, f"Установлен: {target_name}. Изменено: {len(changes)}"
        
    except Exception as e:
        return False, f"Ошибка: {str(e)}"


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
