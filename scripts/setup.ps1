# Первичная настройка SK-Reporter (Windows). Запуск из корня репозитория.
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

if (-not (Test-Path "webapp\main.py")) {
    Write-Error "Запустите скрипт из клона sk-reporter (нужен webapp\main.py)."
}

if (Test-Path "webapp\venv") {
    Write-Host "Удаляю устаревший webapp\venv..."
    Remove-Item -Recurse -Force "webapp\venv"
}

Write-Host "Создаю venv в корне..."
python -m venv venv
& ".\venv\Scripts\Activate.ps1"
pip install --upgrade pip
pip install -e .

Write-Host ""
Write-Host "Готово. Запуск:"
Write-Host "  .\venv\Scripts\Activate.ps1"
Write-Host "  cd webapp"
Write-Host "  python -m uvicorn main:app --reload --host 127.0.0.1 --port 8000"
Write-Host ""
