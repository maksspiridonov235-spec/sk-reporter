# SK-Reporter

Объединение ежедневных отчётов строительного контроля.

- Запуск: **[docs/RUN_SERVER.md](docs/RUN_SERVER.md)** (Windows: [ярлык SK-Reporter.bat](docs/RUN_SERVER.md#офисный-пк-ярлык-sk-reporterbat))
- Для сотрудников: **[docs/ДЛЯ_СОТРУДНИКОВ.md](docs/ДЛЯ_СОТРУДНИКОВ.md)**

## Быстрый старт

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -e .
ollama pull gemma4:31b-cloud

cd webapp
python -m uvicorn main:app --reload --host 127.0.0.1 --port 8000
```

Браузер: http://localhost:8000

## Структура

```
sk-reporter/
├── pyproject.toml            # зависимости и пакет sk_reporter
├── sk_reporter/              # ядро: docx, компании, AI-агенты
│   ├── companies.py
│   ├── docx_processing.py
│   ├── template_layout.py
│   ├── paths.py
│   └── agent/
├── data/templates/           # болванки .docx (в git)
├── webapp/                   # FastAPI: только HTTP и UI
│   ├── main.py
│   ├── static/
│   └── templates/
├── scripts/setup.sh          # первичная настройка (macOS/Linux)
├── scripts/setup.ps1         # первичная настройка (Windows)
├── launcher/SK-Reporter.bat  # ярлык для офисного ПC
# рабочие файлы — temp sk_reports_work/ (uploads, results), не в репо
```

## Сценарий

1. Загрузить отчёты → выбрать дату → **Подготовить** → **Сформировать**
2. Скачать готовые файлы справа или ZIP

Шаблоны болванок — в `data/templates/`, в UI не загружаются.

Новая компания — строка в `sk_reporter/companies.py` + болванка `Компания.docx` в `data/templates/`.
