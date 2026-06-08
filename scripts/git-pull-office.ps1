# Безопасный git pull на офисном ПК.
# Болванки docx и project.yaml/luvr/personnel правятся на месте — без skip-worktree pull блокируется.
#
# Использование (из корня репо):
#   .\scripts\git-pull-office.ps1           # pull + ff-only
#   .\scripts\git-pull-office.ps1 -MarkOnly # только пометить локальные файлы
#   .\scripts\git-pull-office.ps1 -ShowLocal # что помечено skip-worktree
param(
    [switch]$MarkOnly,
    [switch]$ShowLocal,
    [switch]$Quiet
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

function Write-Info([string]$Msg) {
    if (-not $Quiet) { Write-Host $Msg }
}

function Get-OfficeLocalPatterns() {
    return @(
        "data/templates/*.docx",
        "data/projects/*/project.yaml",
        "data/luvr/luvr.yaml",
        "data/personnel/personnel.yaml"
    )
}

function Get-OfficeLocalTrackedFiles() {
    $files = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($pattern in Get-OfficeLocalPatterns) {
        $matched = git ls-files $pattern 2>$null
        if ($LASTEXITCODE -ne 0) { continue }
        foreach ($line in $matched) {
            if ($line) { [void]$files.Add($line.Trim()) }
        }
    }
    return @($files)
}

function Enable-OfficeLocalSkipWorktree {
    $files = Get-OfficeLocalTrackedFiles
    if ($files.Count -eq 0) {
        Write-Info "[INFO] Нет tracked-файлов для skip-worktree."
        return
    }
    foreach ($path in $files) {
        git update-index --skip-worktree $path
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "Не удалось skip-worktree: $path"
        }
    }
    Write-Info "[INFO] skip-worktree: $($files.Count) файлов (болванки, project.yaml, luvr, personnel)."
}

function Show-OfficeLocalSkipWorktree {
    git ls-files -v | Select-String "^[a-zS]" | ForEach-Object {
        $line = $_.Line
        if ($line -match "^S\s+(.+)$") { $Matches[1] }
    }
}

if ($ShowLocal) {
    Write-Host "Файлы с skip-worktree:"
    Show-OfficeLocalSkipWorktree
    exit 0
}

Enable-OfficeLocalSkipWorktree

if ($MarkOnly) {
    Write-Info "[INFO] Готово. Теперь: git pull --ff-only"
    exit 0
}

Write-Info "[INFO] git fetch origin..."
git fetch origin
if ($LASTEXITCODE -ne 0) {
    Write-Warning "[WARN] git fetch не удался — сервер запустится на текущем коде."
    exit 1
}

$branch = (git rev-parse --abbrev-ref HEAD).Trim()
if ($branch -ne "main") {
    Write-Warning "[WARN] Ветка $branch, не main — pull пропущен."
    exit 0
}

Write-Info "[INFO] git merge --ff-only origin/main..."
git merge --ff-only origin/main
if ($LASTEXITCODE -ne 0) {
    Write-Warning @"
[WARN] fast-forward не удался (конфликт или локальные коммиты).
На офисном ПК обычно достаточно:
  git fetch origin
  git reset --hard origin/main
После reset код обновится; локальные project.yaml/docx на диске сохранятся (skip-worktree).
"@
    exit 1
}

$head = (git rev-parse --short HEAD).Trim()
Write-Info "[INFO] Код обновлён: $head"
