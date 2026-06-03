ЗАПУСК
=====

См. docs/RUN_SERVER.md и CLEANUP_PLAN.md в корне репозитория.

Кратко (venv в корне):

  python -m venv venv
  pip install -r requirements.txt    # из корня

  cd webapp
  python -m uvicorn main:app --reload --port 8000

  http://localhost:8000


РАБОТА
======

1. Загрузить отчёты (.docx) — левая панель
2. Выбрать дату → Подготовить → Сформировать
3. Скачать из «Готовые файлы» или ZIP

Шаблоны болванок лежат на сервере в data/templates/ — в UI не загружаются.
См. docs/DATA_TEMPLATES.md


НОВАЯ КОМПАНИЯ
==============

Файл companies.py в корне — добавить строку в COMPANIES.
