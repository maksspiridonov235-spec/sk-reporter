# SK-Reporter: объединение отчётов строительного контроля

> План уборки: **[CLEANUP_PLAN.md](CLEANUP_PLAN.md)**  
> Запуск сервера: **[docs/RUN_SERVER.md](docs/RUN_SERVER.md)** (venv в корне `venv/`)  
> Для сотрудников: **[docs/ДЛЯ_СОТРУДНИКОВ.md](docs/ДЛЯ_СОТРУДНИКОВ.md)**

Веб-приложение на **FastAPI** для подготовки, проверки и сборки ежедневных отчётов СК в сводные документы по компаниям.

## Функции

- **Загрузка отчётов** — drag & drop `.docx` / `.doc` в браузере
- **Подготовка** — правки внутри загруженных файлов и обновление даты в болванках на сервере
- **Проверка и исправление** — AI-анализ описаний через Ollama (`qwen3.5:cloud`)
- **Сборка** — склейка отчётов по компаниям в готовые файлы
- **Лента операций** — прогресс и подробный лог каждой операции

Шаблоны болванок лежат на сервере в `contractor_report/болванки (...)/` — **в UI не загружаются**.

## Структура проекта

```
sk-reporter/
├── CLEANUP_PLAN.md
├── requirements.txt          # все Python-зависимости (корень)
├── venv/                     # виртуальное окружение (не в git)
├── docs/                     # инструкции
├── launcher/                 # SK-Reporter.bat (Windows)
├── companies.py              # список компаний и ключевых слов
├── webapp/
│   ├── main.py               # FastAPI app, mount static, routers
│   ├── config.py             # пути, AGENT_ENABLED, Jinja2
│   ├── helpers.py            # общие хелперы (SSE, merge, даты)
│   ├── routes/               # pages, reports, check, prepare, merge, downloads
│   ├── docx_processing.py    # логика .docx
│   ├── static/               # css, js
│   └── templates/index.html
├── agent/                    # Ollama-агенты
├── contractor_report/        # болванки .docx
└── output/                   # исправленные отчёты (*_исправлен.docx)
```

## Установка и запуск

1. **Python 3.9+**, **Ollama** — [https://ollama.com](https://ollama.com)
2. Из корня репозитория:

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
ollama pull qwen3.5:cloud
```

3. Запуск:

```powershell
cd webapp
python -m uvicorn main:app --reload --host 127.0.0.1 --port 8000
```

Браузер: **http://localhost:8000**

Подробнее — в [docs/RUN_SERVER.md](docs/RUN_SERVER.md).

## Сценарий работы

1. Загрузите отчёты инженеров (.docx)
2. Выберите **дату в отчёте**
3. При необходимости — **Проверить и исправить**
4. **Подготовить** → **Сформировать**
5. Скачайте готовые файлы справа или **Скачать (ZIP)**

## Список компаний

Единый источник — **`companies.py`** в корне репозитория. Добавление новой компании:

```python
("Название", ["ключевое слово 1", "ключевое слово 2"]),
```

## API (основное)

| Метод | URL | Назначение |
|-------|-----|------------|
| GET | `/` | Веб-интерфейс |
| POST | `/upload/reports` | Загрузка отчётов |
| GET | `/files/reports` | Список загруженных |
| POST | `/check/descriptions/stream` | AI-проверка (SSE) |
| POST | `/macro/prepare` | Подготовка отчётов (`{"date":"YYYY-MM-DD"}`) |
| POST | `/rename/templates` | Дата в тексте болванок |
| GET | `/merge/all/stream` | Сборка всех компаний (SSE) |
| POST | `/rename/results` | Переименование готовых |
| POST | `/switch-leader-ai/{leader}` | Смена руководителя |
| GET | `/results` | Список готовых файлов |
| GET | `/download/{filename}` | Скачать файл |
| GET | `/download/all.zip` | ZIP всех готовых |
| GET | `/download/fixed/all.zip` | ZIP исправленных |
| DELETE | `/clear/reports`, `/clear/all` | Очистка |

**Только для разработки** (нет кнопки в UI): `GET /diagnose/reports` — диагностика сетки таблиц в загруженных отчётах.

## AI и Ollama

- Модель по умолчанию: **`qwen3.5:cloud`** (см. `agent/ocr_agent.py`)
- При старте сервера в логе: `[INFO] AI agent connected` или `[WARNING] Agent not found`
- В шапке UI: **AI ✓** / **AI выкл**
- Без Ollama сборка работает по ключевым словам в имени файла; AI-проверка описаний недоступна

## Отладка

- Логи сервера — в терминале, где запущен uvicorn
- Консоль браузера — F12 → Network / Console
- Исправленные файлы — папка `output/` в корне репозитория

## Лицензия

Внутреннее корпоративное приложение.
