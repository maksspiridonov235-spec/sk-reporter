@echo off
REM Замените REPO_ROOT на путь к sk-reporter на офисном ПC
set REPO_ROOT=%~dp0..
cd /d "%REPO_ROOT%\webapp"
start "" http://localhost:8000
call "%REPO_ROOT%\venv\Scripts\activate.bat"
python -m uvicorn main:app --host 127.0.0.1 --port 8000
pause
