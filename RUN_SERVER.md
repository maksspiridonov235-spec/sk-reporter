# Запуск сервера

Инструкция перенесена в **[docs/RUN_SERVER.md](docs/RUN_SERVER.md)**.

Кратко: venv в **корне** (`venv/`), сервер из `webapp/`:

```powershell
.\venv\Scripts\Activate.ps1
cd webapp
python -m uvicorn main:app --reload --host 127.0.0.1 --port 8000
```

План уборки проекта: **[CLEANUP_PLAN.md](CLEANUP_PLAN.md)**
