# Один перезапуск: убить 8000, очистить pycache, проверить шаблоны, старт.
$root = Split-Path -Parent $PSScriptRoot
$webapp = Join-Path $root "webapp"
$templates = Join-Path $webapp "templates"

Get-NetTCPConnection -LocalPort 8000 -ErrorAction SilentlyContinue |
    ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }

Remove-Item -Recurse -Force (Join-Path $webapp "__pycache__") -ErrorAction SilentlyContinue

Write-Host "[INFO] git:" (git -C $root rev-parse --short HEAD)
Write-Host "[INFO] templates:"
Get-ChildItem $templates -Filter *.html | ForEach-Object { Write-Host "  $($_.Name) $($_.Length) bytes" }

Set-Location $webapp
& (Join-Path $root "venv\Scripts\python.exe") -m uvicorn main:app --reload --host 127.0.0.1 --port 8000
