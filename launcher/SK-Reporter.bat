@echo off
REM Запуск SK-Reporter на офисном ПК. Корень репо = launcher\..
setlocal EnableDelayedExpansion
set "REPO_ROOT=%~dp0.."

if not exist "%REPO_ROOT%\venv\Scripts\activate.bat" (
  echo [ERROR] Нет venv. Из PowerShell в корне проекта:
  echo   .\scripts\setup.ps1
  pause
  exit /b 1
)

call "%REPO_ROOT%\venv\Scripts\activate.bat"

REM Техэксперт: data\local\te_expert.env (скопировать с te_expert.env.example)
set "TE_ENV=%REPO_ROOT%\data\local\te_expert.env"
if exist "%TE_ENV%" (
  echo [INFO] Загрузка TE_EXPERT из data\local\te_expert.env
  for /f "usebackq eol=# tokens=1,* delims==" %%a in ("%TE_ENV%") do (
    if not "%%a"=="" set "%%a=%%b"
  )
)

REM Корпоративный прокси/SSL: set SK_REPORTER_PIP_TRUSTED=1 перед запуском bat
set "PIP_EXTRA="
if "%SK_REPORTER_PIP_TRUSTED%"=="1" (
  set "PIP_EXTRA=--trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org"
  echo [INFO] SK_REPORTER_PIP_TRUSTED=1
)

echo [INFO] Обновление зависимостей pip install -e ...
pip install --default-timeout=180 %PIP_EXTRA% -e "%REPO_ROOT%"
if errorlevel 1 (
  echo.
  echo [ERROR] pip install не удался. Часто помогает:
  echo   set SK_REPORTER_PIP_TRUSTED=1
  echo   затем снова SK-Reporter.bat
  echo Подробнее: docs\RUN_SERVER.md
  pause
  exit /b 1
)

python -c "import openpyxl, yaml, xlrd" 2>nul
if errorlevel 1 (
  echo [WARN] openpyxl, yaml или xlrd не найдены — доустановка...
  pip install --default-timeout=180 %PIP_EXTRA% "openpyxl>=3.1,<4" "xlrd>=2.0,<3" "pyyaml>=6.0,<7"
  if errorlevel 1 (
    echo [ERROR] Не удалось установить openpyxl/pyyaml. Нужен интернет до PyPI.
    pause
    exit /b 1
  )
)

for /f %%i in ('git -C "%REPO_ROOT%" rev-parse --short HEAD 2^>nul') do set GIT_HEAD=%%i
if not defined GIT_HEAD set GIT_HEAD=unknown

REM Офис: pull не блокируется болванками docx и project.yaml (skip-worktree)
powershell -NoProfile -ExecutionPolicy Bypass -File "%REPO_ROOT%\scripts\git-pull-office.ps1" -Quiet 2>nul
if errorlevel 1 (
  echo [WARN] git pull не выполнен — работаем на текущем коде ^(!GIT_HEAD!^)
) else (
  for /f %%i in ('git -C "%REPO_ROOT%" rev-parse --short HEAD 2^>nul') do set GIT_HEAD=%%i
)

echo [INFO] SK-Reporter git: !GIT_HEAD!
echo [INFO] После git pull перезапустите этот bat и Ctrl+F5 в браузере.

cd /d "%REPO_ROOT%\webapp"
start "" http://localhost:8000
python -m uvicorn main:app --host 127.0.0.1 --port 8000
pause
