<#
.SYNOPSIS
Try a workbook by hand against the fake Jira, without the automatic tests.

.DESCRIPTION
1. Stops a fake Jira that these scripts started earlier on that port.
2. Builds out\release\jira2dashboard.xlsm if it is missing (when no -Workbook is given).
3. Writes the copy out\inter\jira2dashboard_mock.xlsm, with Config, Periods and PIs filled for the fake Jira
   (.\scripts\run.ps1 mock-workbook). The workbook given is not changed.
4. Starts the fake Jira, 1200 issues and the known anomalies, in a minimized window (start-mock.ps1).
5. Opens the copy in a new Excel of its own, so that a crash cannot take other workbooks with it. Then click
   "Read the Jira configuration", then "Extract from Jira and compute".
Stop the fake Jira with .\scripts\stop-mock.ps1. Close the copy before running this script again.
ASCII only: PowerShell 5.1 reads a file without BOM as ANSI.

.PARAMETER Workbook
The workbook to try. Default out\release\jira2dashboard.xlsm (built if missing).

.PARAMETER Port
Port of the fake Jira. Default 8099.

.EXAMPLE
.\scripts\manual-test.ps1
Tries the workbook of out\release.

.EXAMPLE
.\scripts\manual-test.ps1 -Workbook C:\Downloads\jira2dashboard.xlsm -Port 8100
Tries another version of the workbook, with a fake Jira on port 8100.
.PARAMETER Help
Show this help. --help and -? do the same.
#>
[CmdletBinding(PositionalBinding = $false)]
param(
    [string]$Workbook = "",
    [int]$Port = 8099,
    [switch]$Help,
    [Parameter(ValueFromRemainingArguments = $true)][string[]]$Rest = @()
)
if ($Help -or $Rest -contains "--help") { Get-Help $PSCommandPath -Detailed; exit 0 }
if ($Rest.Count -gt 0) { Write-Host "manual-test: unknown argument $($Rest -join ' '). See .\scripts\manual-test.ps1 --help"; exit 2 }
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
