# Запуск сервера SK-Reporter

Откройте терминал **в корне репозитория** — там, где лежат папки `webapp`, `agent`, `contractor_report` и файл `README.md`.

Проверка:

```bash
ls webapp/main.py agent/requirements.txt README.md
```

Если все три есть — вы в нужном месте.

---

## macOS — первый раз

```bash
cd webapp
python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
pip install -r ../agent/requirements.txt
```

Ollama (для «Проверить и исправить», руководителя, AI-сборки): https://ollama.com — установить, запустить, затем:

```bash
ollama pull qwen3.5:cloud
ollama pull gemma4:31b-cloud
```

## macOS — каждый запуск

```bash
cd webapp
source venv/bin/activate
python3 -m uvicorn main:app --reload --host 127.0.0.1 --port 8000
```

Браузер: http://localhost:8000  
Остановка: Ctrl+C

---

## Windows — первый раз

```powershell
cd webapp
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install --upgrade pip
pip install -r requirements.txt
pip install -r ..\agent\requirements.txt
```

## Windows — каждый запуск

```powershell
cd webapp
.\venv\Scripts\Activate.ps1
python -m uvicorn main:app --reload --host 127.0.0.1 --port 8000
```

---

## Что должно быть в логе при старте

```
[INFO] Templates dir: ... (21 шаблонов)
[INFO] AI agent connected: qwen3.5:cloud via Ollama
INFO:     Uvicorn running on http://127.0.0.1:8000
```

Шаблоны подрядчика уже в `contractor_report/болванки (шаблоны не вырезать только копировать)` — в UI их загружать не нужно.

Исправленные отчёты после проверки: папка `output/` в корне репозитория (`*_исправлен.docx`).

---

## Ошибки

| Ошибка | Причина |
|--------|---------|
| `requirements.txt` не найден | Вы не в `webapp` — сначала `cd webapp` из корня репозитория |
| `Папка с болванками не найдена` | Запуск не из полного клона (нет `contractor_report`) |
| `[WARNING] Agent not found` | Не выполнен `pip install -r ../agent/requirements.txt` |

В Cursor: Terminal → New Terminal — обычно уже открывается в корне проекта.
