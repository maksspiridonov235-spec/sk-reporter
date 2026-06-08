@echo off
REM Офис: обновить код, не трогая болванки и project.yaml. Запуск из launcher\ или корня.
setlocal
set "REPO_ROOT=%~dp0.."
cd /d "%REPO_ROOT%"

echo [INFO] SK-Reporter: office git update...
powershell -NoProfile -ExecutionPolicy Bypass -File "%REPO_ROOT%\scripts\git-pull-office.ps1"
if errorlevel 1 (
  echo.
  echo [ERROR] Не удалось обновить. Если reset падает — один раз в PowerShell:
  echo   git update-index --no-skip-worktree data/projects/SVA-WLL-K058-002-DD-00-AS_00.4/project.yaml
  echo   git update-index --no-skip-worktree data/projects/SUP-WLL-K084-003-DD-00-TX_00.2/project.yaml
  echo   git fetch origin
  echo   git reset --hard origin/main
  pause
  exit /b 1
)
echo [INFO] Готово. Перезапустите SK-Reporter.bat
pause
