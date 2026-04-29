param(
    [string]$SkillDir = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"

$launcher = Join-Path $SkillDir "scripts\md2word_gui.bat"
if (-not (Test-Path -LiteralPath $launcher)) {
    throw "md2word_gui.bat not found: $launcher"
}

$desktop = [Environment]::GetFolderPath("Desktop")
$sendTo = Join-Path $env:APPDATA "Microsoft\Windows\SendTo"

if (-not (Test-Path -LiteralPath $sendTo)) {
    New-Item -ItemType Directory -Path $sendTo -Force | Out-Null
}

function New-Md2WordShortcut {
    param(
        [string]$Path,
        [string]$Target
    )

    $ws = New-Object -ComObject WScript.Shell
    $shortcut = $ws.CreateShortcut($Path)
    $shortcut.TargetPath = $Target
    $shortcut.WorkingDirectory = $env:USERPROFILE
    $shortcut.Description = "Markdown to Word"
    $shortcut.Save()
}

Write-Host "======================================"
Write-Host "  md2word shortcut install/update"
Write-Host "======================================"

$desktopShortcut = Join-Path $desktop "md2word.lnk"
$sendToShortcutName = "md2word " + [char]0x8F6C + "Word.lnk"
$sendToShortcut = Join-Path $sendTo $sendToShortcutName

New-Md2WordShortcut -Path $desktopShortcut -Target $launcher
Write-Host "[OK] Desktop shortcut: $desktopShortcut"

New-Md2WordShortcut -Path $sendToShortcut -Target $launcher
Write-Host "[OK] SendTo shortcut: $sendToShortcut"

Get-ChildItem -LiteralPath $sendTo -Filter "md2word*.lnk" | ForEach-Object {
    if ($_.FullName -ne $sendToShortcut) {
        Remove-Item -LiteralPath $_.FullName -Force
        Write-Host "[OK] Removed duplicate SendTo shortcut: $($_.FullName)"
    }
}

Write-Host ""
Write-Host "Done. Re-run install_or_update.ps1 after updating md2word to refresh shortcuts."
