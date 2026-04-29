param(
    [string]$RepoUrl = "https://github.com/wenhaoa/md2word-pandoc.git",
    [string]$InstallDir = "$env:USERPROFILE\.codex\skills\md2word-pandoc"
)

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "======================================"
Write-Host "  md2word install/update"
Write-Host "======================================"
Write-Host ""

$parentDir = Split-Path -Parent $InstallDir
if (-not (Test-Path -LiteralPath $parentDir)) {
    New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
}

if (Test-Path -LiteralPath (Join-Path $InstallDir ".git")) {
    Write-Host "[1/2] Check remote updates: $InstallDir"
    git -C $InstallDir remote set-url origin $RepoUrl
    git -C $InstallDir fetch origin main

    $local = (git -C $InstallDir rev-parse HEAD).Trim()
    $remote = (git -C $InstallDir rev-parse origin/main).Trim()

    if ($local -eq $remote) {
        Write-Host "      Already up to date: $local"
    } else {
        $dirty = git -C $InstallDir status --porcelain
        if ($dirty) {
            throw "Remote update found, but local uncommitted changes exist. Commit, back up, or clean local changes before updating."
        }
        Write-Host "      Remote update found, updating: $local -> $remote"
        git -C $InstallDir pull --ff-only origin main
    }
} elseif (Test-Path -LiteralPath $InstallDir) {
    $items = Get-ChildItem -LiteralPath $InstallDir -Force
    if ($items.Count -gt 0) {
        throw "Install directory exists but is not a Git repository: $InstallDir. Back it up or remove it before retrying."
    }
    Write-Host "[1/2] Clone skill: $InstallDir"
    git clone $RepoUrl $InstallDir
} else {
    Write-Host "[1/2] Clone skill: $InstallDir"
    git clone $RepoUrl $InstallDir
}

$shortcutInstaller = Join-Path $InstallDir "scripts\install_shortcuts.ps1"
if (-not (Test-Path -LiteralPath $shortcutInstaller)) {
    throw "Shortcut installer not found: $shortcutInstaller"
}

Write-Host ""
Write-Host "[2/2] Refresh Desktop and SendTo shortcuts"
& powershell -NoProfile -ExecutionPolicy Bypass -File $shortcutInstaller

Write-Host ""
Write-Host "Done. Re-run this script after updating md2word to refresh shortcuts."
