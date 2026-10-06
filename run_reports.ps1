param([switch]$InventoryOnly)
$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
try {
    $reportPython = $null
    $reportPythonArgs = @()
    # Use the installed Codex runtime when available on this computer.
    $bundledPython = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
    if (Test-Path -LiteralPath $bundledPython) {
        $reportPython = $bundledPython
    } elseif (Get-Command py -ErrorAction SilentlyContinue) {
        $reportPython = 'py'
        $reportPythonArgs = @('-3')
    } elseif (Get-Command python -ErrorAction SilentlyContinue) {
        $reportPython = 'python'
    } else {
        throw 'Python was not found. Follow the one-time setup in README.md.'
    }
    & $reportPython @reportPythonArgs -c 'import openpyxl'
    if ($LASTEXITCODE -ne 0) {
        throw 'The openpyxl package is missing. Follow the one-time setup in README.md.'
    }
    if ($InventoryOnly) {
        Write-Host '55+ Daily Inventory Update (local JSON only; no GitHub publication)'
        $inventoryDate = Read-Host 'Inventory update date (YYYY-MM-DD), or press Enter for today'
        $inventoryArgs = @((Join-Path $PSScriptRoot 'inventory_update.py'))
        if ($inventoryDate.Trim()) { $inventoryArgs += @('--inventory-date', $inventoryDate.Trim()) }
        & $reportPython @reportPythonArgs @inventoryArgs
        exit $LASTEXITCODE
    }
    Write-Host '55+ Community Market Reports'
    Write-Host 'Close the two exports in Excel before continuing.'
    Write-Host 'Report date excludes that day from closed sales. Example: 2026-10-01 covers 2025-10-01 through 2026-09-30.'
    $reportDate = Read-Host 'Report date (YYYY-MM-DD), or press Enter for today'
    $reportArgs = @((Join-Path $PSScriptRoot 'market_report.py'))
    if ($reportDate.Trim()) { $reportArgs += @('--as-of', $reportDate.Trim()) }
    $inventoryDate = Read-Host 'Inventory source refresh date (YYYY-MM-DD), or Enter to preserve an unchanged source date / use today for a new source'
    if ($inventoryDate.Trim()) { $reportArgs += @('--inventory-date', $inventoryDate.Trim()) }
    & $reportPython @reportPythonArgs @reportArgs
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    Write-Host 'Finished. Open Output\market_summary.xlsx. Historical reports and source exports are saved in Archive.'
} catch {
    Write-Host ('Unable to create report: ' + $_.Exception.Message) -ForegroundColor Red
    exit 1
}
