# Офисный ПК: обновить код с GitHub, сохранить локальные data (болванки, yaml).
#   .\scripts\git-pull-office.ps1
#   .\scripts\git-pull-office.ps1 -MarkOnly
param(
    [switch]$MarkOnly,
    [switch]$Quiet
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root
git config core.quotepath false 2>$null | Out-Null

function Write-Info([string]$Msg) {
    if (-not $Quiet) { Write-Host $Msg }
}

function Clear-AllSkipWorktree {
    $n = 0
    git -c core.quotepath=false ls-files -v | ForEach-Object {
        if ($_ -match '^[sS]\s+(.+)$') {
            git update-index --no-skip-worktree -- $Matches[1].Trim()
            if ($LASTEXITCODE -eq 0) { $script:n++ }
        }
    }
    if ($n -gt 0) { Write-Info "[INFO] Снято skip-worktree: $n файлов." }
}

function Test-GitTracked([string]$RelativePath) {
    git ls-files --error-unmatch -- $RelativePath 2>$null | Out-Null
    return ($LASTEXITCODE -eq 0)
}

function Get-OfficeLocalTrackedFiles() {
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

    foreach ($rel in @("data/luvr/luvr.yaml")) {
        $full = Join-Path $Root ($rel -replace "/", [IO.Path]::DirectorySeparatorChar)
        if ((Test-Path -LiteralPath $full) -and (Test-GitTracked $rel)) {
            [void]$files.Add($rel)
        }
    }
    return @($files)
}

function Enable-OfficeLocalSkipWorktree {
    $files = Get-OfficeLocalTrackedFiles
    foreach ($path in $files) {
        git update-index --skip-worktree -- $path 2>$null | Out-Null
    }
    Write-Info "[INFO] skip-worktree включён: $($files.Count) файлов."
}

function Backup-OfficeLocalFiles {
    param([string[]]$Files, [string]$BackupDir)
    New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null
    $manifest = New-Object System.Collections.Generic.List[string]
    $i = 0
    foreach ($rel in $Files) {
        $src = Join-Path $Root ($rel -replace "/", [IO.Path]::DirectorySeparatorChar)
        if (-not (Test-Path -LiteralPath $src)) { continue }
        $destName = ("{0:D4}__{1}" -f $i, ($rel -replace '[\\/:*?"<>|]', '_'))
        Copy-Item -LiteralPath $src -Destination (Join-Path $BackupDir $destName) -Force
        $manifest.Add("$destName`t$rel")
        $i++
    }
    [IO.File]::WriteAllLines(
        (Join-Path $BackupDir "manifest.tsv"),
        $manifest,
        [Text.UTF8Encoding]::new($false)
    )
    return $i
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
        $dest = Join-Path $Root ($parts[1] -replace "/", [IO.Path]::DirectorySeparatorChar)
        $dir = Split-Path -LiteralPath $dest -Parent
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
        Copy-Item -LiteralPath (Join-Path $BackupDir $parts[0]) -Destination $dest -Force
        $restored++
    }
    return $restored
}

Clear-AllSkipWorktree

if ($MarkOnly) {
    Enable-OfficeLocalSkipWorktree
    Write-Info "[INFO] Готово."
    exit 0
}

$localFiles = Get-OfficeLocalTrackedFiles
$backupDir = Join-Path $env:TEMP ("sk-reporter-pull-{0}" -f (Get-Date -Format "yyyyMMddHHmmss"))
$backedUp = Backup-OfficeLocalFiles -Files $localFiles -BackupDir $backupDir
Write-Info "[INFO] Backup data: $backedUp файлов."

Clear-AllSkipWorktree

Write-Info "[INFO] git fetch origin..."
git fetch origin
if ($LASTEXITCODE -ne 0) { exit 1 }

Write-Info "[INFO] git reset --hard origin/main..."
git reset --hard origin/main
if ($LASTEXITCODE -ne 0) {
    Write-Warning "[WARN] reset не удался. См. docs/RUN_SERVER.md"
    exit 1
}

Restore-OfficeLocalFiles -BackupDir $backupDir | Out-Null
Remove-Item -LiteralPath $backupDir -Recurse -Force -ErrorAction SilentlyContinue

Enable-OfficeLocalSkipWorktree
Write-Info "[INFO] Код: $(git rev-parse --short HEAD)"
