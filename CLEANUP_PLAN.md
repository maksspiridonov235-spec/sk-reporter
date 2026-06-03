# План уборки SK-Reporter

Документ фиксирует цель, структуру и порядок работ. Возвращайтесь к нему при паузах между фазами.

**Статус:** фаза 4 завершена — план уборки выполнен  
**Окружение Python:** `venv/` в **корне** репозитория (не `webapp/venv`)

---

## Зачем

- Разделить UI, API и утилиты — проще сопровождать на одном офисном ПC.
- Убрать мусор и устаревшие инструкции.
- Не ломать рабочий сценарий: загрузка → дата → проверка → подготовить → сформировать → скачать.

---

## Целевая структура (после фаз 0–3)

```
sk-reporter/
├── CLEANUP_PLAN.md          ← этот файл
├── requirements.txt         ← все Python-зависимости
├── venv/                    ← виртуальное окружение (не в git)
├── docs/                    ← инструкции
├── launcher/                ← SK-Reporter.bat для Windows
├── archive/                 ← не используется в проде
├── scripts/                 ← diagnose и утилиты
├── data/templates/          ← болванки .docx (в git)
├── output/                  ← исправленные docx (не в git)
├── webapp/
│   ├── main.py
│   ├── config.py
│   ├── helpers.py
│   ├── routes/
│   ├── docx_processing.py
│   ├── static/              ← фаза 1: css, js
│   └── templates/
│       └── index.html
├── agent/                   ← Ollama-агенты
├── companies.py
├── apply_template_layout.py
└── test_data/
```

---

## Карта фронта → API (не удалять при рефакторинге)

| UI | Эндпоинты |
|----|-----------|
| Загрузка отчётов | `POST /upload/reports`, `GET /files/reports` |
| Дата в отчёте | `POST /rename/templates`, `POST /macro/prepare`, `POST /rename/results` |
| Проверить и исправить | `POST /check/descriptions/stream` |
| Руководитель | `POST /switch-leader-ai/{leader}` |
| Подготовить | `POST /macro/prepare` |
| Сформировать | `GET /merge/all/stream` |
| Скачать ZIP | `GET /download/all.zip`, `GET /download/fixed/all.zip` |
| Очистить / сброс | `DELETE /clear/reports`, `/clear/results`, `/clear/all` |

**Только для отладки (нет кнопки в UI):** `GET /diagnose/reports`

---

## Фаза 0 — инвентаризация и фундамент

- [x] Создать `CLEANUP_PLAN.md`
- [x] Папка `archive/` + перенос мусора
- [x] `docs/RUN_SERVER.md` — venv в корне
- [x] `docs/ДЛЯ_СОТРУДНИКОВ.md` + `launcher/SK-Reporter.bat`
- [x] `requirements.txt` в корне (webapp + agent)
- [x] `scripts/setup.sh` / `scripts/setup.ps1`
- [x] Обновить `.gitignore`, `.vscode/launch.json`
- [ ] **Вручную на офисном ПC:** пересоздать venv в корне, удалить `webapp/venv`

### В archive/

| Было | Зачем архив |
|------|-------------|
| `agent/inject_agent.py.backup` | Старая копия |
| `analyze_doc.py`, `analyze_template.py` | Разовые скрипты анализа |
| `tmp/patch.py` | Черновой патч |
| `.webapp/` | Ошибочная копия конфига |

---

## Фаза 1 — фронтенд

- [x] `webapp/static/css/app.css` — весь CSS из `index.html`
- [x] `webapp/static/js/` — модули по зонам (api, activity, reports, check, prepare, merge, downloads, help)
- [x] `StaticFiles` в `main.py`
- [x] `index.html` — только разметка + `<script src="/static/...">`
- [ ] Прогон сценария на офисном ПC

---

## Фаза 2 — мёртвое на бэке

- [x] Убрать неиспользуемые `companies` / `agent_enabled` из шаблона или показать статус AI в UI
- [x] Пометить `/diagnose/reports` как dev-only
- [x] Обновить `README.md`, убрать устаревшее (`Flask`, загрузка шаблонов в UI, Qwen 7b)
- [x] `companies.py` — единственный источник списка компаний (не `docx_processing.py`)

---

## Фаза 3 — бэкенд по папкам

- [x] `webapp/routes/` — reports, check, merge, downloads (+ pages, prepare)
- [x] Тонкий `main.py` (mount static, include routers)
- [x] (Опционально) `src/core/`, `src/docx/` — не делали; достаточно `config.py` + `helpers.py`

**Не делать в фазе 3:** полный пакет `pyproject.toml` / pip install -e — лишний риск для одного ПC.

---

## Фаза 4 — данные (отдельное решение)

- [x] `contractor_report/болванки (...)/` → `data/templates/`
- [x] Политика git для `.docx` шаблонов (в репо; остальные `.docx` игнорируются)
- [x] `README.md` в `data/templates/` + `docs/DATA_TEMPLATES.md`

---

## Миграция venv (webapp → корень)

**Один раз на машине разработки:**

```powershell
# Windows — из корня репозитория
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

После проверки запуска удалить старое: `webapp\venv\`

```bash
# macOS
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
rm -rf webapp/venv
```

---

## Критерий «фаза завершена»

| Фаза | Проверка |
|------|----------|
| 0 | Сервер стартует из корневого venv, docs актуальны |
| 1 | UI выглядит и работает как до разрезания |
| 2 | README совпадает с UI, нет вводящих в заблуждение путей |
| 3 | API те же URL, `main.py` короче и читаемее |
