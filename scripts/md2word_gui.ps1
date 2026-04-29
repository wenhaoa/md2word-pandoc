param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$InputFiles
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$conversionScript = Join-Path $scriptDir "run_conversion.js"

function Wait-And-Exit {
    param(
        [int]$Code,
        [int]$Seconds = 10
    )

    if ($Code -eq 0) {
        Start-Sleep -Seconds $Seconds
    } else {
        Write-Host ""
        Read-Host "Press Enter to close"
    }
    exit $Code
}

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    Write-Host "[Error] Node.js not found. Install it first: winget install OpenJS.NodeJS.LTS"
    Wait-And-Exit -Code 1
}

if (-not (Test-Path -LiteralPath $conversionScript)) {
    Write-Host "[Error] conversion script not found: $conversionScript"
    Wait-And-Exit -Code 1
}

$mdFile = $null
if ($InputFiles -and $InputFiles.Count -gt 0) {
    $mdFile = $InputFiles[0]
} else {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Select Markdown file"
    $dialog.Filter = "Markdown files (*.md)|*.md|All files (*.*)|*.*"
    $dialog.FilterIndex = 1
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $mdFile = $dialog.FileName
    }
}

if (-not $mdFile) {
    Write-Host "Canceled."
    Wait-And-Exit -Code 0 -Seconds 2
}

$mdFile = [System.IO.Path]::GetFullPath($mdFile)
if (-not (Test-Path -LiteralPath $mdFile)) {
    Write-Host "[Error] file not found: $mdFile"
    Wait-And-Exit -Code 1
}

Write-Host ""
Write-Host "Converting: $mdFile"
Write-Host "====================================="
Write-Host ""

& node $conversionScript $mdFile --open
$code = $LASTEXITCODE

if ($code -ne 0) {
    Write-Host ""
    Write-Host "====================================="
    Write-Host "Conversion failed. Check the messages above."
    Wait-And-Exit -Code $code
}

Write-Host "====================================="
Write-Host "Conversion complete. Word file has been opened."
Wait-And-Exit -Code 0 -Seconds 10
