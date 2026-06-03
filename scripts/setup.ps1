# Первичная настройка SK-Reporter (Windows). Запуск из корня репозитория.
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

if (-not (Test-Path "webapp\main.py")) {
    Write-Error "Запустите скрипт из клона sk-reporter (нужен webapp\main.py)."
}

function Invoke-Pip {
    param(
        [Parameter(ValueFromRemainingArguments = $true)]
        [string[]]$PipArgs
    )
    $allArgs = @("--default-timeout", "180") + $PipArgs
    if ($env:SK_REPORTER_PIP_TRUSTED -eq "1") {
        $allArgs += @(
            "--trusted-host", "pypi.org",
            "--trusted-host", "pypi.python.org",
            "--trusted-host", "files.pythonhosted.org"
        )
        Write-Host "Режим SK_REPORTER_PIP_TRUSTED=1 (корпоративный прокси/SSL)."
    }
    & pip @allArgs
    if ($LASTEXITCODE -ne 0) {
        Write-Error @"
pip завершился с ошибкой (код $LASTEXITCODE).

Частые причины:
  • getaddrinfo failed — нет интернета/DNS (проверьте ping pypi.org; иногда нужен VPN офиса)
  • SSL certificate — `$env:SK_REPORTER_PIP_TRUSTED = "1"` и снова setup
  • venv занят — закройте SK-Reporter.bat и все окна с (venv)

Подробнее: docs/RUN_SERVER.md
"@
    }
}

function Remove-VenvSafe {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return }
    Write-Host "Удаляю $Path ..."
    try {
        Remove-Item -Recurse -Force $Path -ErrorAction Stop
    } catch {
        Write-Error @"
Не удалось удалить $Path (файлы заняты).

Закройте:
  • окно SK-Reporter.bat (чёрная консоль)
  • все PowerShell с префиксом (venv)
  • Cursor/VS Code терминалы в этом проекте

Затем в новом PowerShell (без venv):
  deactivate
  Remove-Item -Recurse -Force venv
  .\scripts\setup.ps1
"@
    }
}

# Нельзя пересоздать venv, пока он активен — python.exe заблокирован
if ($env:VIRTUAL_ENV) {
    Write-Host "Деактивирую текущий venv..."
    if (Get-Command deactivate -ErrorAction SilentlyContinue) {
        deactivate
    }
}

Remove-VenvSafe "webapp\venv"
Remove-VenvSafe "venv"

Write-Host "Создаю venv в корне..."
python -m venv venv
& ".\venv\Scripts\Activate.ps1"

Write-Host "Проверка доступа к PyPI..."
$dnsOk = $false
try {
    $null = Resolve-DnsName pypi.org -ErrorAction Stop
    $dnsOk = $true
} catch {
    Write-Warning "DNS не резолвит pypi.org — pip, скорее всего, не скачает пакеты. Нужен рабочий интернет (иногда только с VPN офиса)."
}

Write-Host "Обновление pip (необязательно, пропускаем при ошибке)..."
& pip install --default-timeout 60 --upgrade pip 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Warning "pip не обновлён — продолжаю с версией в venv."
}

Write-Host "Устанавливаю sk-reporter..."
Invoke-Pip @("install", "-e", ".")

Write-Host ""
Write-Host "Готово. Запуск:"
Write-Host "  .\venv\Scripts\Activate.ps1"
Write-Host "  cd webapp"
Write-Host "  python -m uvicorn main:app --reload --host 127.0.0.1 --port 8000"
Write-Host ""
Write-Host "Или ярлык: launcher\SK-Reporter.bat"
Write-Host ""
