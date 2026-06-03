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

function Get-PythonCommand {
    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) {
        return @($python.Source)
    }

    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) {
        return @($py.Source, "-3")
    }

    return $null
}

function Invoke-Python {
    param(
        [string[]]$PythonCommand,
        [string[]]$Arguments
    )

    if ($PythonCommand.Count -gt 1) {
        & $PythonCommand[0] $PythonCommand[1..($PythonCommand.Count - 1)] @Arguments
    } else {
        & $PythonCommand[0] @Arguments
    }
}

function Test-PythonModule {
    param(
        [string[]]$PythonCommand,
        [string]$ModuleName
    )

    Invoke-Python -PythonCommand $PythonCommand -Arguments @("-c", "import $ModuleName") *> $null
    return $LASTEXITCODE -eq 0
}

function Ensure-PythonDependencies {
    $pythonCommand = Get-PythonCommand
    if (-not $pythonCommand) {
        Write-Host "[Error] Python not found. Install it first: winget install Python.Python.3.12"
        Wait-And-Exit -Code 1
    }

    $missingPackages = @()
    if (-not (Test-PythonModule -PythonCommand $pythonCommand -ModuleName "docx")) {
        $missingPackages += "python-docx"
    }
    if (-not (Test-PythonModule -PythonCommand $pythonCommand -ModuleName "docxcompose")) {
        $missingPackages += "docxcompose"
    }

    if ($missingPackages.Count -eq 0) {
        return
    }

    Write-Host "[Info] Installing missing Python packages: $($missingPackages -join ', ')"
    Invoke-Python -PythonCommand $pythonCommand -Arguments (@("-m", "pip", "install", "--user") + $missingPackages)
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[Error] Failed to install Python packages."
        Write-Host "        Please run: python -m pip install --user $($missingPackages -join ' ')"
        Wait-And-Exit -Code 1
    }
}

Ensure-PythonDependencies

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
