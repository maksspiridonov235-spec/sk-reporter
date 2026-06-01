# Запуск сервера SK-Reporter

Нужен **Python 3.9+**. Шаблоны подрядчиков уже лежат в репозитории (`contractor_report/болванки ...`) — отдельно загружать их в браузер не нужно.

Для функций с ИИ (проверка, руководитель, определение компании) нужен **Ollama**: https://ollama.com — установите и запустите приложение, затем подтяните модели, которые указаны в `agent/*.py` (сейчас в коде: `qwen3.5:cloud`, `gemma4:31b-cloud`). Без Ollama сервер всё равно стартует, но «умные» кнопки не сработают.

---

## macOS (первый запуск)

### 1. Терминал

Откройте **Terminal** (Программы → Утилиты → Terminal) или встроенный терминал в Cursor.

### 2. Перейдите в папку проекта

Подставьте свой путь к клону репозитория:

```bash
cd ~/путь/к/sk-reporter
```

Пример, если проект на рабочем столе:

```bash
cd ~/Desktop/sk-reporter
```

### 3. Python (если ещё нет)

Проверка:

```bash
python3 --version
```

Если команды нет или версия ниже 3.9:

```bash
brew install python@3.12
```

(нужен [Homebrew](https://brew.sh): `/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"`)

### 4. Виртуальное окружение (один раз)

```bash
cd webapp
python3 -m venv venv
source venv/bin/activate
```

В начале строки должно появиться `(venv)`.

Установка зависимостей:

```bash
pip install --upgrade pip
pip install -r requirements.txt
pip install -r ../agent/requirements.txt
```

### 5. Ollama (для ИИ, один раз)

1. Скачайте и установите Ollama с https://ollama.com  
2. Запустите Ollama (иконка в меню — сервис должен работать).  
3. В терминале (можно вне venv):

```bash
ollama pull qwen3.5:cloud
ollama pull gemma4:31b-cloud
```

Если какая-то cloud-модель недоступна — смотрите ошибки в терминале при нажатии «Проверить и исправить»; без моделей merge по ключевым словам и макросы могут работать.

### 6. Запуск сервера

Из папки `webapp`, с активным `(venv)`:

```bash
cd ~/путь/к/sk-reporter/webapp
source venv/bin/activate
python3 -m uvicorn main:app --reload --host 127.0.0.1 --port 8000
```

Другой порт, например 3000:

```bash
python3 -m uvicorn main:app --reload --host 127.0.0.1 --port 3000
```

Ожидаемые строки в терминале:

```
[INFO] Templates dir: ... (21 шаблонов)
[INFO] AI agent connected: qwen3.5:cloud via Ollama
INFO:     Uvicorn running on http://127.0.0.1:8000
```

Если видите `[WARNING] Agent not found` — переустановите зависимости агента (шаг 4).

### 7. Браузер

Откройте: **http://localhost:8000** (или тот порт, который указали).

### 8. Остановка

В том же терминале: **Ctrl + C**

### Повторный запуск (уже настраивали venv)

```bash
cd ~/путь/к/sk-reporter/webapp
source venv/bin/activate
python3 -m uvicorn main:app --reload --host 127.0.0.1 --port 8000
```

---

## Windows (PowerShell)

### 1. PowerShell

### 2. Папка проекта

```powershell
cd C:\Users\Anton\Desktop\sk-reporter\webapp
```

(свой путь к репозиторию)

### 3. Виртуальное окружение (первый раз)

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install --upgrade pip
pip install -r requirements.txt
pip install -r ..\agent\requirements.txt
```

### 4. Запуск

```powershell
.\venv\Scripts\Activate.ps1
python -m uvicorn main:app --reload --host 127.0.0.1 --port 8000
```

Браузер: **http://localhost:8000**

Остановка: **Ctrl + C**

---

## Использование (актуально для UI)

1. **Загрузите отчёты** — зона «Отчёты от инженеров» (`.docx` / `.doc`).
2. При необходимости: **Проверить и исправить**, **Руководитель**, **Макросы**.
3. **Сформировать все отчёты** — сборка в болванки из `contractor_report`.
4. **Забрать всё (ZIP)** или скачать из «Готовые файлы».

Исправленные отчёты после проверки сохраняются в папку `output/` в корне проекта (`*_исправлен.docx`).

Временные загрузки: системная temp-папка, подпапка `sk_reports_work` (на Mac обычно `/var/folders/...` или `$TMPDIR`).

---

## Если что-то не работает

| Симптом | Что проверить |
|--------|----------------|
| `command not found: python3` | Установить Python через brew или с python.org |
| `Папка с болванками не найдена` | Запускать из полного клона репозитория, не только папку `webapp` |
| ИИ не отвечает / долго висит | Ollama запущен? `ollama list` в терминале |
| «Шаблон не найден» на «Применить размеры шаблона» | В болванках нет файла `Ежедневный отчет Шаблон.docx` — известная проблема |
| Порт занят | Сменить `--port 8001` или закрыть другой uvicorn |

Логи смотрите в терминале, где запущен uvicorn.
