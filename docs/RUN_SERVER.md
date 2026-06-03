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

**Проще всего — ярлык** (см. раздел [«Ярлык SK-Reporter.bat»](#офисный-пк-ярлык-sk-reporterbat) ниже).

Вручную из терминала:

```powershell
.\venv\Scripts\Activate.ps1
cd webapp
python -m uvicorn main:app --reload --host 127.0.0.1 --port 8000
```

Браузер: http://localhost:8000  
Остановка: Ctrl+C

---

## Офисный ПК: ярлык SK-Reporter.bat

Файл **`launcher/SK-Reporter.bat`** — запуск сервера одним двойным щелчком (без ручного терминала каждый день).

### Первый раз на этом компьютере

1. Клонировать или обновить репозиторий (`git pull`).
2. В PowerShell из **корня** `sk-reporter`:

```powershell
.\scripts\setup.ps1
```

3. Установить [Ollama](https://ollama.com), запустить, затем:

```powershell
ollama pull gemma4:31b-cloud
```

### Ярлык на рабочем столе

1. Проводник → папка `sk-reporter\launcher\`.
2. Правой кнопкой по **`SK-Reporter.bat`** → **Отправить** → **Рабочий стол (создать ярлык)**.
3. Переименовать ярлык в **SK-Reporter**.

Или создать ярлык вручную: объект  
`C:\Users\ИМЯ\Desktop\sk-reporter\launcher\SK-Reporter.bat`  
(подставьте свой путь к проекту).

**Важно:** ярлык должен указывать на `.bat` **внутри** `launcher\`, а не на копию батника elsewhere. Батник сам находит корень проекта (`launcher\..`).

### Каждый рабочий день

1. Дважды щёлкнуть ярлык **SK-Reporter**.
2. Откроется браузер: http://localhost:8000
3. **Чёрное окно консоли не закрывать** — там работает сервер.
4. Закончили работу — закрыть это окно (или Ctrl+C).

### Если окно сразу закрывается

Запустить из PowerShell, чтобы увидеть ошибку:

```powershell
cd C:\Users\ИМЯ\Desktop\sk-reporter\launcher
.\SK-Reporter.bat
```

| Ошибка | Что сделать |
|--------|-------------|
| Нет `venv` | Из корня проекта: `.\scripts\setup.ps1` |
| `python` не найден | Установить Python с [python.org](https://www.python.org/downloads/), галочка «Add python.exe to PATH» |
| Порт 8000 занят | Закрыть старое окно SK-Reporter или перезагрузить ПК |

Сотрудникам без терминала: **[docs/ДЛЯ_СОТРУДНИКОВ.md](ДЛЯ_СОТРУДНИКОВ.md)**.

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
