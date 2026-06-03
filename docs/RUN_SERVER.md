# Запуск сервера SK-Reporter

Терминал открывайте **в корне репозитория** — там, где `sk_reporter/`, `webapp/`, `pyproject.toml`.

Проверка:

```bash
ls webapp/main.py pyproject.toml data/templates
```

---

## Первый раз (Windows)

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install --upgrade pip
pip install -e .
```

Или: `.\scripts\setup.ps1`

Ollama: https://ollama.com — установить, запустить, затем:

```powershell
ollama pull gemma4:31b-cloud
```

---

## Каждый запуск (Windows)

```powershell
.\venv\Scripts\Activate.ps1
cd webapp
python -m uvicorn main:app --reload --host 127.0.0.1 --port 8000
```

Браузер: http://localhost:8000  
Остановка: Ctrl+C

Для сотрудников: ярлык `launcher/SK-Reporter.bat` (см. `docs/ДЛЯ_СОТРУДНИКОВ.md`).

---

## Первый раз (macOS)

```bash
python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip
pip install -e .
```

Или: `./scripts/setup.sh`

```bash
ollama pull gemma4:31b-cloud
```

---

## Каждый запуск (macOS)

```bash
source venv/bin/activate
cd webapp
python3 -m uvicorn main:app --reload --host 127.0.0.1 --port 8000
```

---

## Лог при старте

```
[INFO] Templates dir: ... (N шаблонов)
[INFO] AI agent connected: gemma4:31b-cloud via Ollama
INFO:     Uvicorn running on http://127.0.0.1:8000
```

Шаблоны: **`data/templates/`** — в UI не загружаются.

Исправленные отчёты: `output/` в корне (`*_исправлен.docx`).

---

## Отладка (dev)

Эндпоинт **`GET /diagnose/reports`** — только для разработки, кнопки в UI нет. Проверяет сетку таблиц в загруженных отчётах. Ответ содержит `"dev_only": true`.

Пример (после загрузки отчётов в UI):

```bash
curl http://127.0.0.1:8000/diagnose/reports
```

---

## Ошибки

| Ошибка | Причина |
|--------|---------|
| `pyproject.toml` не найден | Терминал не в корне репозитория |
| `Папка с болванками не найдена` | Нет `data/templates/` или пустая — скопируйте болванки или сделайте `git pull` |
| `[WARNING] Agent not found` | Не выполнен `pip install -e .` в корневом venv |
