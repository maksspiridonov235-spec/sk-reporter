# SK-Reporter

Веб-приложение для **строительного контроля**: ежедневные отчёты (.docx), проверка предписаний (Excel), планирование (PostgreSQL), отчёт инженера по ВОР.

- Запуск prod: **https://sk-reporter.relaxdev.ru** — см. **[docs/ДЛЯ_СОТРУДНИКОВ.md](docs/ДЛЯ_СОТРУДНИКОВ.md)**
- Локальная разработка: **[docs/RUN_SERVER.md](docs/RUN_SERVER.md)**
- Для сотрудников: **[docs/ДЛЯ_СОТРУДНИКОВ.md](docs/ДЛЯ_СОТРУДНИКОВ.md)**
- Для разработки / AI: **[AGENTS.md](AGENTS.md)**, **[docs/memory.md](docs/memory.md)**

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

## Структура репозитория

Папка на диске **`sk-reporter/`** — весь проект. Внутри **`sk_reporter/`** (с подчёркиванием) — Python-пакет с логикой: в имени модуля нельзя дефис, поэтому так и задумано, это не дубликат.

```
sk-reporter/                         ← корень репозитория (git, docs, данные)
│
├── webapp/                          ← сайт: то, что открывается в браузере
│   ├── main.py                      ← FastAPI-сервер, маршруты /api/*
│   ├── templates/                   ← HTML (главная, planning, daily, …)
│   └── static/                      ← JS, CSS
│
├── sk_reporter/                     ← «мозги» на Python (import sk_reporter.*)
│   ├── agent/                       ← Ollama: проверка описаний в docx
│   ├── docx_processing.py           ← подготовка и сборка ежедневных отчётов
│   ├── companies.py                 ← список подрядчиков и болванок
│   ├── luvr_store.py                ← ЛУВР (yaml ↔ xlsx)
│   ├── planning_data.py             ← API данных для раздела «Планирование»
│   ├── deployment_store.py          ← расстановка из xlsm
│   ├── appendix7_store.py           ← Приложение 7
│   ├── prescriptions/               ← проверка Excel предписаний
│   └── engineer/                    ← код «Инженер ФИО» (ВОР, сборка docx)
│
├── data/                            ← данные офиса (часть в git, часть локально)
│   ├── templates/                   ← болванки подрядчиков (.docx)
│   ├── projects/                    ← проекты, ВОР, назначения инженеров
│   ├── personnel/                   ← исходный Excel для импорта в PostgreSQL
│   ├── luvr/                        ← ЛУВР, luvr.yaml, шаблоны xlsm
│   └── tk/                          ← каталог технологических карт
│
├── scripts/                         ← утилиты из терминала (setup, сборка yaml)
├── docs/                            ← инструкции, memory.md, контекст продукта
└── pyproject.toml                   ← зависимости; pip install -e .
```

**Рабочие файлы пользователей** (загруженные отчёты, результаты) — во временной папке `sk_reports_work/` или системном temp (`…/sk_reports_work/uploads/`), **не в репозитории**.

### Что за что на главной странице

| Раздел в UI | Маршруты | Код | Данные |
|-------------|----------|-----|--------|
| **Планирование** | `/planning` | `planning_data.py`, `personnel_db.py`, `otkk_db.py` | PostgreSQL `personnel`, `otkk_cards` |
| **Отчётность** | `/reporting` → `/daily`, `/prescriptions` | `sk_reporter/agent/`, `docx_processing.py`, `prescriptions/` | `data/templates/`, temp uploads |
| **Инженер ФИО** | `/engineer-hub` → `/engineer/{person_id}` | `sk_reporter/engineer/` | PostgreSQL `personnel` |

### Частые вопросы

- **`sk_reporter/engineer/`** — Python-модуль отчёта инженера (ВОР, docx); ФИО из PostgreSQL.
- **`scripts/`** — не второй проект, а редкие команды (`setup.sh`, `build_engineer_data.py --luvr`).
- **Тяжёлые xlsx/xlsm** — класть в `data/luvr/` (или `data/prescriptions/`), не в корень репо и не коммитить без необходимости (см. `.gitignore`).

## Сценарий: ежедневные отчёты

1. Отчётность → Ежедневные отчёты → загрузить .docx → дата и руководитель → **Проверить и сформировать**
2. Скачать готовые файлы справа или ZIP

Шаблоны болванок — в `data/templates/`, в UI не загружаются.

Новая компания — строка в `sk_reporter/companies.py` + болванка `Компания.docx` в `data/templates/`.
