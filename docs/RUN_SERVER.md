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
| `SSL: CERTIFICATE_VERIFY_FAILED` / `self-signed certificate` при `setup.ps1` | Корпоративный прокси. В PowerShell **перед** setup: `$env:SK_REPORTER_PIP_TRUSTED = "1"`, затем снова `.\scripts\setup.ps1` |
| `Read timed out` при `pip install` | Медленная сеть. Повторить setup; при необходимости `$env:SK_REPORTER_PIP_TRUSTED = "1"`. Или вручную: `.\venv\Scripts\Activate.ps1` → `pip install --default-timeout=180 -e .` |
| Скрипт написал «Готово», но `pip install` падал | Зависимости не установились — проверьте: `pip show sk-reporter`. Если пусто — повторите setup с `SK_REPORTER_PIP_TRUSTED=1` |

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
[INFO] Ollama: mode=local|cloud, host=..., model=gemma4:31b-cloud, api_key=yes|no, ping=ok|fail
INFO:     Uvicorn running on http://127.0.0.1:8000
```

### Ollama: офис vs облако (Relax Dev)

| Режим | Env |
|-------|-----|
| **Офис** (локальный Ollama) | ничего, или `OLLAMA_HOST=http://127.0.0.1:11434`; `ollama signin`, `ollama pull gemma4:31b-cloud` |
| **Облако** (без daemon на сервере) | `OLLAMA_API_KEY` — ключ с [ollama.com/settings/keys](https://ollama.com/settings/keys); `OLLAMA_HOST=https://ollama.com`; опционально `OLLAMA_MODEL=gemma4:31b-cloud` |

Ключи **не коммитить**. При `ping=fail` в логе — AI (отчёты, предписания) не заработает.

Проверка ключа:

```bash
curl -s https://ollama.com/api/tags -H "Authorization: Bearer $OLLAMA_API_KEY"
```

Шаблоны: **`data/templates/`** — в UI не загружаются.

Загруженные и исправленные отчёты: temp `sk_reports_work/uploads/` (см. `UPLOAD_DIR` в `webapp/main.py`). Скачивание с суффиксом `_исправлен` — через UI, файл на диске с исходным именем.

---

## Отладка (dev)

Эндпоинт **`GET /diagnose/reports`** — только для разработки, кнопки в UI нет. Проверяет сетку таблиц в загруженных отчётах. Ответ содержит `"dev_only": true`.

Пример (после загрузки отчётов в UI):

```bash
curl http://127.0.0.1:8000/diagnose/reports
```

---

## Офисный ПК: git pull без «грязного дерева»

На сервере **ежедневно меняются** файлы, которые лежат в git:

| Путь | Почему «modified» |
|------|-------------------|
| `data/templates/*.docx` | сборка отчётов переименовывает/трогает болванки |
| `data/projects/*/project.yaml` | назначения инженеров в «Проекты» |
| `data/luvr/luvr.yaml` | отметки ЛУВР на сайте |
| `data/personnel/personnel.yaml` | справочник персонала |

Cursor/VS Code и `git pull` блокируются, пока эти изменения не убраны — **это нормально для офиса**, данные нужны на месте.

**Решение:** скрипт `scripts/git-pull-office.ps1`:
1. Резервная копия локальных data (болванки, yaml).
2. `git reset --hard origin/main` — код с GitHub без merge-конфликтов.
3. Восстановление data с диска + **`skip-worktree`**.

**Важно:** `git stash -- data/` **не убирает** файлы с уже включённым `skip-worktree` — pull всё равно падает. Не используйте stash для обновления; только скрипт или блок ниже.

**Разблокировка прямо сейчас** (Anton / если `reset --hard` пишет *not uptodate* — виноват **skip-worktree**):

```powershell
cd C:\Users\Anton\Desktop\sk-reporter

git update-index --no-skip-worktree data/projects/SVA-WLL-K058-002-DD-00-AS_00.4/project.yaml
git update-index --no-skip-worktree data/projects/SUP-WLL-K084-003-DD-00-TX_00.2/project.yaml

git fetch origin
git reset --hard origin/main

git log -1 --oneline
```

Должно показать свежий коммит (не `961e927`). Потом:

```powershell
.\scripts\git-pull-office.ps1 -MarkOnly
```

Если скрипта ещё нет — после `reset --hard` он появится, команда сработает.

**Полный цикл с backup** (если нужно):

```powershell
$bak = "$env:TEMP\sk-reporter-backup"
Copy-Item -Recurse -Force data\templates $bak\templates -ErrorAction SilentlyContinue
Copy-Item -Recurse -Force data\projects $bak\projects -ErrorAction SilentlyContinue

git update-index --no-skip-worktree data/projects/SVA-WLL-K058-002-DD-00-AS_00.4/project.yaml
git update-index --no-skip-worktree data/projects/SUP-WLL-K084-003-DD-00-TX_00.2/project.yaml
git fetch origin
git reset --hard origin/main

Copy-Item -Recurse -Force $bak\templates\* data\templates\ -ErrorAction SilentlyContinue
Copy-Item -Recurse -Force $bak\projects\* data\projects\ -ErrorAction SilentlyContinue
.\scripts\git-pull-office.ps1 -MarkOnly
```

Ярлык **`launcher\office-git-update.bat`** — то же самое двойным щелчком (после первого успешного reset).

**Обычное обновление:**

```powershell
.\scripts\git-pull-office.ps1
```

На **Mac (разработка)** skip-worktree **не включать** — иначе не увидите изменения в `git status`.

---

## Предписания: доступ к Техэксперт

Проверка предписаний (`/prescriptions`) ищет нормативный документ из ячейки **B19** в **Техэксперт** (или в интернете при fallback) и сверяет фрагмент с текстом замечания (**B18**). В исправленный Excel в **B19** записывается **краткий заголовок** из выдачи TE (например «Приказ Ростехнадзора от 11.12.2020 N 519») и пункты инженера — не развёрнутое название из текста документа и не текст модели.

**Рекомендуется (офис):** скопировать `data/local/te_expert.env.example` → `data/local/te_expert.env`, вписать логин и пароль. Файл в `.gitignore`; `SK-Reporter.bat` и сервер подхватывают его автоматически.

Альтернатива — переменные окружения перед запуском (пароль **не** хранить в репозитории):

| Переменная | Пример | Назначение |
|------------|--------|------------|
| `TE_EXPERT_BASE_URL` | `http://248960.te-cloud.ru` | Адрес облачного сервера |
| `TE_EXPERT_CATALOG` | `/docs` | Виртуальный каталог (по умолчанию `/docs`) |
| `TE_EXPERT_LOGIN` | *(логин)* | Учётная запись Техэксперт |
| `TE_EXPERT_PASSWORD` | *(пароль)* | Пароль |
| `TE_EXPERT_USE_BROWSER` | `0` | `0` — HTTP API (рекомендуется), `1` — запасной Playwright |
| `TE_EXPERT_INTERNET_FALLBACK` | `1` | `1` — искать НД в интернете, если Техэксперт недоступен; `0` — только Техэксперт |

**Windows (PowerShell, перед bat):**

```powershell
$env:TE_EXPERT_BASE_URL = "http://248960.te-cloud.ru"
$env:TE_EXPERT_LOGIN = "ваш_логин"
$env:TE_EXPERT_PASSWORD = "ваш_пароль"
```

**macOS / Linux:**

```bash
export TE_EXPERT_BASE_URL="http://248960.te-cloud.ru"
export TE_EXPERT_LOGIN="ваш_логин"
export TE_EXPERT_PASSWORD="ваш_пароль"
```

Для te-cloud достаточно HTTP API; Playwright нужен только как запасной вариант (`TE_EXPERT_USE_BROWSER=1`):

```bash
pip install "sk-reporter[browser]"
playwright install chromium
```

Если вход в Техэксперт не проходит — проверьте логин/пароль в браузере на том же URL. В отчёте проверки будет предупреждение «нормативка не сверена онлайн».

**Проверка с Mac / перед деплоем:**

```bash
export TE_EXPERT_BASE_URL="http://248960.te-cloud.ru"
export TE_EXPERT_LOGIN="..."
export TE_EXPERT_PASSWORD="..."
python scripts/test_techexpert.py "Приказ 519, п. 44"
```

| Симптом | Что делать |
|---------|------------|
| `toomanyusers` | Закрыть лишние вкладки/сессии TE в браузере; подождать и повторить |
| В отчёте источник «интернет» при открытом TE | TE API вернул ошибку — смотреть `te_fallback_error` в отчёте или лог сервера |
| B19 без полного названия приказа | Ожидаемо: в ячейку пишется **краткий title** из TE + пункты |

---

## Ошибки

| Ошибка | Причина |
|--------|---------|
| `pyproject.toml` не найден | Терминал не в корне репозитория |
| `Папка с болванками не найдена` | Нет `data/templates/` или пустая — скопируйте болванки или сделайте `git pull` |
| `[WARNING] Agent not found` | Не выполнен `pip install -e .` в корневом venv |
| `SSL: CERTIFICATE_VERIFY_FAILED` / timeout при `pip install` | Корпоративная сеть: `$env:SK_REPORTER_PIP_TRUSTED = "1"` и снова `.\scripts\setup.ps1` (см. раздел про ярлык выше) |
| `getaddrinfo failed` / `No matching distribution found` | **Нет интернета до PyPI** (DNS). Проверьте `ping pypi.org`. Без VPN интернет может пропасть — включите VPN офиса + `SK_REPORTER_PIP_TRUSTED=1` |
| `Permission denied` на `venv\Scripts\python.exe` | venv занят. Закройте SK-Reporter.bat и все `(venv)` терминалы, `deactivate`, удалите `venv`, setup снова |
| После `git pull` старый UI или нет `/daily` | `git pull` + перезапуск uvicorn (Ctrl+C → снова run). Проверка: `/health` → `templates_on_disk`: `daily.html`, `home.html`; `has_daily_route`: true |
| `No module named 'openpyxl'` на `/api/planning/luvr` или в Планировании | Код обновился, **venv — нет**. Закрыть bat → в PowerShell: `.\venv\Scripts\Activate.ps1` → `pip install -e .` (при SSL: `$env:SK_REPORTER_PIP_TRUSTED="1"`). Или перезапустить обновлённый `launcher\SK-Reporter.bat` — он сам доустановит openpyxl |
