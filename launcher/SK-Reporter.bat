@echo off
REM Замените REPO_ROOT на путь к sk-reporter на офисном ПC
set REPO_ROOT=%~dp0..
call "%REPO_ROOT%\venv\Scripts\activate.bat"
pip install -q -e "%REPO_ROOT%"
cd /d "%REPO_ROOT%\webapp"
for /f %%i in ('git -C "%REPO_ROOT%" rev-parse --short HEAD 2^>nul') do set GIT_HEAD=%%i
if not defined GIT_HEAD set GIT_HEAD=unknown
echo [INFO] SK-Reporter build: %GIT_HEAD%
echo [INFO] После git pull обязательно перезапустите этот bat и Ctrl+F5 в браузере.
start "" http://localhost:8000
python -m uvicorn main:app --host 127.0.0.1 --port 8000
pause
