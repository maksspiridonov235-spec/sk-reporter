"""
Агент на базе Ollama + Qwen для анализа отчётов СК.
Определяет компанию по содержимому документа, а не по имени файла.
"""

import ollama
from docx import Document
from typing import Optional, List, Dict
import re
import json

# Список известных компаний (синхронизирован с main.py)
KNOWN_COMPANIES = [
    "Евракор", "Лесные технологии", "ЮНС", "НГСК", "Сибитек", "ЭСМ",
    "НГП", "РОСЭКСПО", "ТПС", "ТЭКПРО", "ЮПИ", "УГГ", "ЮГС", "ТВС",
    "НСС", "ОТ и ТБ", "Стройфинансгрупп"
]

SYSTEM_PROMPT = f"""
Ты — эксперт-аналитик документов строительного контроля.
Твоя задача: проанализировать текст отчёта и определить, какой компании он принадлежит.

Известные компании:
{json.dumps(KNOWN_COMPANIES, ensure_ascii=False)}

Инструкции:
1. Внимательно прочитай текст. Ищи названия компаний в шапке, подписях, таблицах.
2. Учитывай сокращения, аббревиатуры и полные названия (ООО, АО, ПАО).
3. Если найдено несколько компаний, выбери ту, чей отчёт представлен (обычно исполнитель).
4. Верни ТОЛЬКО название компании из списка выше.
5. Если компания не найдена или не входит в список, верни строку "UNKNOWN".

Формат ответа: Строка с названием компании или "UNKNOWN". Никакого лишнего текста.
"""

def extract_text_from_docx(filepath: str, max_pages: int = 3) -> str:
    """
    Извлекает текст из первых N страниц документа.
    Обычно вся нужная информация (название компании) находится в начале.
    """
    try:
        doc = Document(filepath)
        text_parts = []
        
        # Считаем параграфы (грубая оценка страниц: ~30-40 строк на страницу)
        lines_count = 0
        max_lines = max_pages * 35
        
        for para in doc.paragraphs:
            if lines_count >= max_lines:
                break
            text = para.text.strip()
            if text:
                text_parts.append(text)
                lines_count += len(text) // 80 + 1  # Грубая оценка строк
        
        # Также добавим текст из таблиц (там часто бывают шапки)
        for table in doc.tables[:2]:  # Первые 2 таблицы
            for row in table.rows[:5]:  # Первые 5 строк таблицы
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    if cell_text and len(cell_text) > 3:
                        text_parts.append(cell_text)
        
        return "\n".join(text_parts)
    
    except Exception as e:
        print(f"Ошибка чтения файла {filepath}: {e}")
        return ""


def detect_company_with_ai(filepath: str, model: str = "qwen3.5:397b-cloud") -> Optional[str]:
    """
    Использует AI для определения компании в документе.
    
    Args:
        filepath: Путь к файлу .docx
        model: Название модели в Ollama
    
    Returns:
        Название компании из списка KNOWN_COMPANIES или None
    """
    # 1. Извлекаем текст
    text = extract_text_from_docx(filepath)
    
    if not text or len(text) < 50:
        print(f"⚠️ Файл {filepath} слишком короткий или пустой для анализа.")
        return None
    
    # 2. Отправляем запрос к Ollama
    try:
        response = ollama.chat(
            model=model,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": f"Проанализируй этот текст отчёта и определи компанию:\n\n{text[:3000]}"} # Ограничиваем контекст
            ],
            options={
                "temperature": 0.1,  # Минимальная случайность для точности
                "num_predict": 50    # Короткий ответ
            }
        )
        
        ai_answer = response['message']['content'].strip()
        
        # 3. Парсим ответ
        # Очищаем от кавычек и лишнего мусора, если модель вдруг добавила пояснения
        clean_name = re.sub(r'["\'\.]', '', ai_answer).strip()
        
        # Проверяем, есть ли ответ в списке известных
        for company in KNOWN_COMPANIES:
            if company.lower() in clean_name.lower() or clean_name.lower() in company.lower():
                print(f"✅ AI определил компанию: {company} (ответ модели: {ai_answer})")
                return company
        
        if clean_name.upper() == "UNKNOWN":
            print(f"⚠️ AI не смог определить компанию для {filepath}")
            return None
            
        # Если модель вернула что-то похожее, но не точное совпадение
        print(f"❓ AI вернул нестандартный ответ: '{ai_answer}'. Попытка сопоставления...")
        return None
        
    except Exception as e:
        print(f"❌ Ошибка при запросе к Ollama: {e}")
        return None


def detect_company_hybrid(filepath: str) -> Optional[str]:
    """
    Гибридный метод: сначала быстрый поиск по ключевым словам (как раньше),
    если не найдено — используем AI.
    Это экономит ресурсы и время.
    """
    from webapp.docx_processing import COMPANIES
    
    filename_lower = filepath.lower()
    
    # Быстрая проверка по имени файла (мгновенно)
    for company_name, keywords in COMPANIES:
        for keyword in keywords:
            if keyword.lower() in filename_lower:
                return company_name
    
    # Если по имени не нашли — подключаем тяжелую артиллерию (AI)
    print(f"🔍 По имени файла компания не найдена. Запускаю AI-анализ для {filepath}...")
    return detect_company_with_ai(filepath)


if __name__ == "__main__":
    # Тестовый запуск
    import sys
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        result = detect_company_hybrid(file_path)
        print(f"Итоговый результат: {result}")
    else:
        print("Укажите путь к файлу для проверки: python agent/ocr_agent.py path/to/file.docx")
