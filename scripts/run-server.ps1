# Один перезапуск: убить 8000/8010, очистить pycache, старт на 8010 (8000 часто занят призраком).
$root = Split-Path -Parent $PSScriptRoot
$webapp = Join-Path $root "webapp"
$templates = Join-Path $webapp "templates"
$port = 8010

foreach ($p in 8000, 8010) {
    Get-NetTCPConnection -LocalPort $p -ErrorAction SilentlyContinue |
        ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }
}

Remove-Item -Recurse -Force (Join-Path $webapp "__pycache__") -ErrorAction SilentlyContinue

Write-Host "[INFO] git:" (git -C $root rev-parse --short HEAD)
Write-Host "[INFO] templates:"
Get-ChildItem $templates -Filter *.html | ForEach-Object { Write-Host "  $($_.Name) $($_.Length) bytes" }
Write-Host "[INFO] Open: http://127.0.0.1:$port/"

Set-Location $webapp
& (Join-Path $root "venv\Scripts\python.exe") -m uvicorn main:app --reload `
  --reload-dir (Join-Path $root "webapp") `
  --reload-dir (Join-Path $root "sk_reporter") `
  --host 127.0.0.1 --port $port
