# Безопасный git pull на офисном ПК.
# Болванки docx и project.yaml/luvr/personnel правятся на месте — обычный pull блокируется.
# Стратегия: backup локальных data → reset --hard origin/main → restore → skip-worktree.
#
# Использование (из корня репо):
#   .\scripts\git-pull-office.ps1            # обновить код, сохранить локальные data
#   .\scripts\git-pull-office.ps1 -MarkOnly  # только skip-worktree
#   .\scripts\git-pull-office.ps1 -ShowLocal # список skip-worktree
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

function Test-GitTracked([string]$RelativePath) {
    git ls-files --error-unmatch -- $RelativePath 2>$null | Out-Null
    return ($LASTEXITCODE -eq 0)
}

function Get-OfficeLocalTrackedFiles() {
    # Через диск, не через вывод git ls-files — иначе PowerShell ломает кириллицу в путях.
    $files = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    $templatesDir = Join-Path $Root "data\templates"
    if (Test-Path -LiteralPath $templatesDir) {
        Get-ChildItem -LiteralPath $templatesDir -Filter *.docx -File | ForEach-Object {
            $rel = "data/templates/$($_.Name)"
            if (Test-GitTracked $rel) { [void]$files.Add($rel) }
        }
    }

    $projectsDir = Join-Path $Root "data\projects"
    if (Test-Path -LiteralPath $projectsDir) {
        Get-ChildItem -LiteralPath $projectsDir -Directory | ForEach-Object {
            $yaml = Join-Path $_.FullName "project.yaml"
            if (Test-Path -LiteralPath $yaml) {
                $rel = "data/projects/$($_.Name)/project.yaml"
                if (Test-GitTracked $rel) { [void]$files.Add($rel) }
            }
        }
    }

    foreach ($rel in @("data/luvr/luvr.yaml", "data/personnel/personnel.yaml")) {
        $full = Join-Path $Root ($rel -replace "/", [IO.Path]::DirectorySeparatorChar)
        if ((Test-Path -LiteralPath $full) -and (Test-GitTracked $rel)) {
            [void]$files.Add($rel)
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
        git update-index --skip-worktree -- $path
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "Не удалось skip-worktree: $path"
        }
    }
    Write-Info "[INFO] skip-worktree: $($files.Count) файлов."
}

function Show-OfficeLocalSkipWorktree {
    git -c core.quotepath=false ls-files -v | ForEach-Object {
        if ($_ -match '^[sS]\s+(.+)$') { $Matches[1] }
    }
}

function Backup-OfficeLocalFiles {
    param(
        [string[]]$Files,
        [string]$BackupDir
    )
    New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null
    $manifest = New-Object System.Collections.Generic.List[string]
    $i = 0
    foreach ($rel in $Files) {
        $src = Join-Path $Root ($rel -replace "/", [IO.Path]::DirectorySeparatorChar)
        if (-not (Test-Path -LiteralPath $src)) { continue }
        $destName = ("{0:D4}__{1}" -f $i, ($rel -replace '[\\/:*?"<>|]', '_'))
        $dest = Join-Path $BackupDir $destName
        Copy-Item -LiteralPath $src -Destination $dest -Force
        $manifest.Add("$destName`t$rel")
        $i++
    }
    $manifestPath = Join-Path $BackupDir "manifest.tsv"
    [IO.File]::WriteAllLines($manifestPath, $manifest, [Text.UTF8Encoding]::new($false))
    return $manifest.Count
}

function Restore-OfficeLocalFiles {
    param([string]$BackupDir)
    $manifestPath = Join-Path $BackupDir "manifest.tsv"
    if (-not (Test-Path -LiteralPath $manifestPath)) { return 0 }
    $restored = 0
    foreach ($line in [IO.File]::ReadAllLines($manifestPath)) {
        if (-not $line.Trim()) { continue }
        $parts = $line -split "`t", 2
        if ($parts.Count -lt 2) { continue }
        $destName, $rel = $parts[0], $parts[1]
        $src = Join-Path $BackupDir $destName
        $dest = Join-Path $Root ($rel -replace "/", [IO.Path]::DirectorySeparatorChar)
        $destDir = Split-Path -LiteralPath $dest -Parent
        if (-not (Test-Path -LiteralPath $destDir)) {
            New-Item -ItemType Directory -Force -Path $destDir | Out-Null
        }
        Copy-Item -LiteralPath $src -Destination $dest -Force
        $restored++
    }
    return $restored
}

if ($ShowLocal) {
    Write-Host "Файлы с skip-worktree:"
    Show-OfficeLocalSkipWorktree
    exit 0
}

# Кириллица в git status на Windows
git config core.quotepath false 2>$null | Out-Null

Enable-OfficeLocalSkipWorktree

if ($MarkOnly) {
    Write-Info "[INFO] Готово. Для обновления кода: .\scripts\git-pull-office.ps1"
    exit 0
}

Write-Info "[INFO] git fetch origin..."
git fetch origin
if ($LASTEXITCODE -ne 0) {
    Write-Warning "[WARN] git fetch не удался."
    exit 1
}

$branch = (git rev-parse --abbrev-ref HEAD).Trim()
if ($branch -ne "main") {
    Write-Warning "[WARN] Ветка $branch, не main — обновление пропущено."
    exit 0
}

$localFiles = Get-OfficeLocalTrackedFiles
$backupDir = Join-Path $env:TEMP ("sk-reporter-pull-{0}" -f (Get-Date -Format "yyyyMMddHHmmss"))
$backedUp = Backup-OfficeLocalFiles -Files $localFiles -BackupDir $backupDir
Write-Info "[INFO] Резервная копия локальных data: $backedUp файлов."

Write-Info "[INFO] git reset --hard origin/main..."
git reset --hard origin/main
if ($LASTEXITCODE -ne 0) {
    Write-Warning "[WARN] reset не удался."
    exit 1
}

$restored = Restore-OfficeLocalFiles -BackupDir $backupDir
Write-Info "[INFO] Восстановлено локальных data: $restored файлов."

try {
    Remove-Item -LiteralPath $backupDir -Recurse -Force -ErrorAction SilentlyContinue
} catch { }

Enable-OfficeLocalSkipWorktree

$head = (git rev-parse --short HEAD).Trim()
Write-Info "[INFO] Код обновлён: $head"
