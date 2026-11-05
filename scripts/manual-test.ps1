# .\scripts\manual-test.ps1 [-Workbook out\release\jira2dashboard.xlsm] [-Port 8099]
# Try a workbook by hand against the fake Jira, without the automatic tests:
#   1. stops a fake Jira that these scripts started earlier on that port;
#   2. builds out\release\jira2dashboard.xlsm if it is missing (when no -Workbook is given);
#   3. writes the copy out\inter\jira2dashboard_mock.xlsm with Config, Periods and PIs filled for the fake Jira
#      (scripts\run.ps1 mock-workbook): the workbook given is not changed;
#   4. starts the fake Jira, 1200 issues and the known anomalies, in a minimized window (scripts\start-mock.ps1);
#   5. opens the copy in a new Excel of its own, so that a crash cannot take other workbooks with it.
# Stop the fake Jira with .\scripts\stop-mock.ps1. ASCII only: PowerShell 5.1 reads a file without BOM as ANSI.
param(
    [string]$Workbook = "",
    [int]$Port = 8099
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

function Fail([string]$message) {
    Write-Host "manual-test: $message"
    exit 1
}

& (Join-Path $PSScriptRoot "stop-mock.ps1") -Port $Port
if (-not $Workbook) {
    $Workbook = Join-Path $root "out\release\jira2dashboard.xlsm"
    if (-not (Test-Path $Workbook)) {
        & (Join-Path $PSScriptRoot "run.ps1") build
        if ($LASTEXITCODE -ne 0) { Fail "the build of the workbook failed (see above)." }
    }
}
& (Join-Path $PSScriptRoot "run.ps1") mock-workbook --workbook $Workbook --port $Port
if ($LASTEXITCODE -ne 0) { Fail "the copy could not be prepared (see above)." }
& (Join-Path $PSScriptRoot "start-mock.ps1") -Port $Port
if ($LASTEXITCODE -ne 0) { Fail "the fake Jira did not start (see above)." }

$copy = Join-Path $root "out\inter\jira2dashboard_mock.xlsm"
Write-Host "manual-test: Excel may crash right after the last search page (known defect). Save your other work first."
Start-Process excel.exe -ArgumentList "/x", "`"$copy`""
Write-Host "manual-test: $copy opens in a new Excel. Enable the content, then click the Dashboard buttons in this order:"
Write-Host "  1. Read the Jira configuration"
Write-Host "  2. Extract from Jira and compute"
Write-Host "manual-test: stop the fake Jira with .\scripts\stop-mock.ps1 -Port $Port"
exit 0
