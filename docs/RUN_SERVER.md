# Запуск сервера SK-Reporter

## Продакшен (RelaxDev) — для сотрудников

**URL:** https://sk-reporter.relaxdev.ru

Сотрудникам ничего запускать не нужно — только браузер. Инструкции: **[docs/ДЛЯ_СОТРУДНИКОВ.md](ДЛЯ_СОТРУДНИКОВ.md)**.

Обновление кода и перезапуск — на стороне RelaxDev (деплой из `main`).

---

## Локальная разработка

Терминал открывайте **в корне репозитория** — там, где `sk_reporter/`, `webapp/`, `pyproject.toml`.

Проверка:

```bash
ls webapp/main.py pyproject.toml data/templates
```

### Первый раз (Windows)

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

### Каждый запуск (Windows)

```powershell
.\venv\Scripts\Activate.ps1
cd webapp
python -m uvicorn main:app --reload --host 127.0.0.1 --port 8000
```

Или из корня: `.\scripts\run-server.ps1` (порт **8010**, убивает занятые 8000/8010).

Браузер: http://localhost:8000 (или :8010)  
Остановка: Ctrl+C

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
[INFO] Ollama: mode=local|cloud, host=..., model=gemma4:31b-cloud, api_key=yes|no, ping=ok|fail
INFO:     Uvicorn running on http://127.0.0.1:8000
```

### Ollama: локально vs облако (RelaxDev)

| Режим | Env |
|-------|-----|
| **Локально** | `OLLAMA_HOST=http://127.0.0.1:11434`; `ollama signin`, `ollama pull gemma4:31b-cloud` |
| **Облако** (без daemon на сервере) | `OLLAMA_API_KEY` — ключ с [ollama.com/settings/keys](https://ollama.com/settings/keys); `OLLAMA_HOST=https://ollama.com`; опционально `OLLAMA_MODEL=gemma4:31b-cloud` |

Ключи **не коммитить**. При `ping=fail` в логе — AI (отчёты, предписания) не заработает.

Проверка ключа:

```bash
curl -s https://ollama.com/api/tags -H "Authorization: Bearer $OLLAMA_API_KEY"
```

Шаблоны: **`data/templates/`** — в UI не загружаются.

Загруженные и исправленные отчёты: temp `sk_reports_work/uploads/` (см. `UPLOAD_DIR` в `webapp/main.py`). Скачивание с суффиксом `_исправлен` — через UI, файл на диске с исходным именем.

---

## Локальный Windows-ПК: git pull без «грязного дерева»

На сервере **ежедневно меняются** файлы, которые лежат в git:

| Путь | Почему «modified» |
|------|-------------------|
| `data/templates/*.docx` | сборка отчётов переименовывает/трогает болванки |
| PostgreSQL `personnel` | справочник сотрудников (RelaxDev) |

Cursor/VS Code и `git pull` блокируются, пока эти изменения не убраны — **это нормально**, если на этом ПК ещё лежат рабочие data.

**Решение:** `git fetch` + `git reset --hard origin/main` (данные планирования — в PostgreSQL на RelaxDev, не в yaml на диске).

**Разблокировка** (если `reset --hard` пишет *not uptodate* из‑за skip-worktree на болванках):

```powershell
cd C:\Users\Anton\Desktop\sk-reporter

git update-index --no-skip-worktree data/templates/*.docx

git fetch origin
git reset --hard origin/main

git log -1 --oneline
```

Должно показать свежий коммит. Болванки при необходимости восстановите из backup или `git checkout -- data/templates/`.

На **Mac (разработка)** skip-worktree **не включать** — иначе не увидите изменения в `git status`.

---

## Предписания: локальная база НД

Проверка (`/prescriptions`) берёт фрагмент нормативки из **`data/normative/`** (`manifest.yaml` + `texts/*.txt`). Техэксперт и интернет **не используются**. B19 в файле **не переписывается** — только B18.

См. `data/normative/README.md` — как добавить документы.

---

## Ошибки

| Ошибка | Причина |
|--------|---------|
| `pyproject.toml` не найден | Терминал не в корне репозитория |
| `Папка с болванками не найдена` | Нет `data/templates/` или пустая — скопируйте болванки или сделайте `git pull` |
| `[WARNING] Agent not found` | Не выполнен `pip install -e .` в корневом venv |
| `SSL: CERTIFICATE_VERIFY_FAILED` / timeout при `pip install` | Корпоративная сеть: `$env:SK_REPORTER_PIP_TRUSTED = "1"` и снова `.\scripts\setup.ps1` |
| `getaddrinfo failed` / `No matching distribution found` | **Нет интернета до PyPI** (DNS). Проверьте `ping pypi.org`. Без VPN интернет может пропасть — включите VPN + `SK_REPORTER_PIP_TRUSTED=1` |
| `Permission denied` на `venv\Scripts\python.exe` | venv занят. Закройте uvicorn и все `(venv)` терминалы, `deactivate`, удалите `venv`, setup снова |
| После `git pull` старый UI | Перезапуск uvicorn (Ctrl+C → снова run). Проверка: `/health` |
| `No module named 'openpyxl'` | `pip install -e .` в активированном venv |
| Расстановка: «LibreOffice не найден» | Linux-сервер: LibreOffice Calc (`soffice` в PATH) или `LIBREOFFICE_PATH` |
