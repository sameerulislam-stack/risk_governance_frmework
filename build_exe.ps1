$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = Join-Path $projectRoot ".venv\Scripts\python.exe"
$tradeWorkbook = Join-Path $projectRoot "data\trade_monitoring.xlsx"

if (-not (Test-Path $python)) {
    throw "Project Python environment not found: $python"
}
if (-not (Test-Path (Join-Path $projectRoot "data\risk.db"))) {
    throw "Equity report database not found."
}
if (-not (Test-Path $tradeWorkbook)) {
    Write-Warning "Trade Monitoring workbook not found. The report will require an upload."
}

Push-Location $projectRoot
try {
    & $python -m pip install -r "requirements.txt"
    if ($LASTEXITCODE -ne 0) {
        throw "Installing the application dependencies failed with exit code $LASTEXITCODE."
    }

    & $python -m pip install -r "requirements-build.txt"
    if ($LASTEXITCODE -ne 0) {
        throw "Installing the executable build dependency failed with exit code $LASTEXITCODE."
    }

    & $python -m PyInstaller --noconfirm --clean "JLI_Risk_Governance.spec"
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller failed with exit code $LASTEXITCODE."
    }

    $releasePath = Join-Path $projectRoot "release"
    $archivePath = Join-Path $releasePath "JLI_Risk_Governance_Windows.zip"
    New-Item -ItemType Directory -Path $releasePath -Force | Out-Null
    if (Test-Path $archivePath) {
        Remove-Item -LiteralPath $archivePath -Force
    }
    Compress-Archive `
        -Path (Join-Path $projectRoot "dist\JLI Risk Governance") `
        -DestinationPath $archivePath `
        -CompressionLevel Optimal
    Write-Output "Built executable folder: $(Join-Path $projectRoot 'dist\JLI Risk Governance')"
    Write-Output "Shareable package: $archivePath"
}
finally {
    Pop-Location
}
