# AI-агент для отчётов СК

Модуль **Ollama + Qwen** для определения компании, сборки, проверки описаний и смены руководителя.

## Как это работает

1. **Быстрый поиск** — ключевые слова в имени файла (список в `companies.py`).
2. **AI-анализ** — если по имени не найдено: текст из документа → Ollama → сопоставление с `companies.py`.
3. **Проверка описаний** — `check_agent.py` + `inject_agent.py` (кнопка «Проверить и исправить» в UI).

## Требования

- **Ollama** — https://ollama.com
- Модель (по умолчанию в коде): **`qwen3.5:cloud`**

```powershell
ollama pull qwen3.5:cloud
ollama list
```

Python-зависимости — из корневого `requirements.txt` (venv в корне репозитория).

## Быстрый тест

```powershell
# из корня, с активированным venv
python agent/ocr_agent.py "путь/к/файлу.docx"
```

## Логи сервера

При старте `webapp/main.py`:

- `[INFO] AI agent connected: qwen3.5:cloud via Ollama` — агент доступен
- `[WARNING] Agent not found` — сборка по ключевым словам, без AI-проверки

В шапке UI: **AI ✓** / **AI выкл**.
