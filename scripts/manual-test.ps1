<#
.SYNOPSIS
Try a workbook by hand against the fake Jira, without the automatic tests.

.DESCRIPTION
1. Stops a fake Jira that these scripts started earlier on that port.
2. Builds out\release\jira2dashboard.xlsm if it is missing (when no -Workbook is given).
3. Writes the copy out\manual_test\jira2dashboard4manualtest.xlsm, with Config, Iterations and PIs filled for the fake
   Jira (.\scripts\run.ps1 mock-workbook). The workbook given is not changed. out\manual_test is never cleaned.
4. Starts the fake Jira, 1200 issues and the known anomalies, in a minimized window (start-mock.ps1).
5. Opens the copy in a new Excel of its own, so that a crash cannot take other workbooks with it. Then click
   "Read the Jira configuration", then "Extract from Jira and compute".
   With -Python: runs the Python extraction on the copy instead (.\scripts\run.ps1 extract, no macro), then opens
   out\manual_test\jira2dashboard.xlsx (values and SVG charts) in a new Excel.
Stop the fake Jira with .\scripts\stop-mock.ps1. Close the copy before running this script again.
ASCII only: PowerShell 5.1 reads a file without BOM as ANSI.
USE CASES
  Click through the macros           .\scripts\manual-test.ps1                       (Read the Jira configuration, then Extract from Jira and compute)
  Try the Python path, no macro      .\scripts\manual-test.ps1 -Python
  Only the scenario art-metrics      .\scripts\manual-test.ps1 -Scenario art-metrics        (own fake Jira on 8101, own copy, only its data)
  The same, nothing to click         .\scripts\manual-test.ps1 -Scenario art-metrics -Run   (extracts, exports the charts, opens their page)
  Try a downloaded version           .\scripts\manual-test.ps1 -Workbook C:\Downloads\jira2dashboard.xlsm
  Two fake Jiras at once             .\scripts\manual-test.ps1 -Port 8100
  The whole chain without clicking   .\scripts\scenario.ps1                          (unattended: Python extraction, checks, clean-up)
  Excel crashed on the last page     a known defect: save your other work first, run it again
  Where is the log                   scripts\logs\manual-test.log (last run only)


.PARAMETER Workbook
The workbook to try. Default out\release\jira2dashboard.xlsm (built if missing).

.PARAMETER Port
Port of the fake Jira. Default 8099, or 8101 with -Scenario.

.PARAMETER Scenario
Try one named scenario of the fake Jira alone, for example art-metrics (25 Features, 4 PIs, 3 teams): a separate instance with only
its own data. The fake Jira serves nothing else (start-mock.ps1 -Scenario), and the copy is out\manual_test\<scenario>\jira2dashboard4<scenario>.xlsm
with its own Projects, Workflows, Transitions, Iterations and PIs. The generic manual test (port 8099, its own copy) is not touched.

.PARAMETER Run
With -Scenario: nothing to click. In one process: the scenario is served, the copy extracts with its macros, the charts are exported
(SVG, PNG) to out\manual_test\charts\<scenario>, the copy is saved with its results, Excel is closed, and the page index.html of the charts opens.

.PARAMETER Python
Try the Python path without VBA instead of the macros: .\scripts\run.ps1 extract on the copy.

.EXAMPLE
.\scripts\manual-test.ps1
Tries the workbook of out\release.

.EXAMPLE
.\scripts\manual-test.ps1 -Workbook C:\Downloads\jira2dashboard.xlsm -Port 8100
Tries another version of the workbook, with a fake Jira on port 8100.

.EXAMPLE
.\scripts\manual-test.ps1 -Scenario art-metrics -Run
Runs only the scenario art-metrics unattended and opens the page of its charts.

.EXAMPLE
.\scripts\manual-test.ps1 -Python
Tries the Python extraction: its jira2dashboard.xlsx opens in a new Excel.

.PARAMETER Help
Show this help. --help and -? do the same.
#>
[CmdletBinding(PositionalBinding = $false)]
param(
    [string]$Workbook = "",
    [int]$Port = 0,
    [string]$Scenario = "",
    [switch]$Run,
    [switch]$Python,
    [switch]$Help,
    [Parameter(ValueFromRemainingArguments = $true)][string[]]$Rest = @()
)
if ($Help -or $Rest -contains "--help") { Get-Help $PSCommandPath -Detailed; exit 0 }
if ($Rest.Count -gt 0) { Write-Host "manual-test: J2D-SCR-001 - unknown argument $($Rest -join ' '). See .\scripts\manual-test.ps1 --help"; exit 2 }
$ErrorActionPreference = "Stop"
$ScenarioPort = 8101                                  # a scenario is its own instance: the generic fake Jira (8099) can run beside it
if ($Port -eq 0) { $Port = if ($Scenario) { $ScenarioPort } else { 8099 } }
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

# Log of this run (last run only): scripts\logs\manual-test.log. A log failure never stops the script.
# Called by another script of this folder that already writes a log ($J2D_LogOwner, seen through the calling scope):
# no second transcript, the lines of this script go to the log of the caller.
if (-not (Get-Variable J2D_LogOwner -ErrorAction SilentlyContinue)) {
    try {
        $logDir = Join-Path $PSScriptRoot "logs"
        New-Item -ItemType Directory -Force $logDir | Out-Null
        Start-Transcript -Path (Join-Path $logDir "manual-test.log") -Force | Out-Null
        $J2D_LogOwner = "manual-test"
    } catch { }
}

function Fail([string]$code, [string]$message) {
    Write-Host "manual-test: $code - $message"
    exit 1
}

if ($Run -and -not $Scenario) { Fail "J2D-SCR-052" "-Run needs -Scenario: the unattended run is the one of a scenario, for example -Scenario art-metrics." }
& (Join-Path $PSScriptRoot "stop-mock.ps1") -Port $Port
if (-not $Workbook) {
    $Workbook = Join-Path $root "out\release\jira2dashboard.xlsm"
    if (-not (Test-Path $Workbook)) {
        & (Join-Path $PSScriptRoot "run.ps1") build
        if ($LASTEXITCODE -ne 0) { Fail "J2D-SCR-016" "the build of the workbook failed (see above)." }
    }
}
$copyArguments = @("--workbook", $Workbook, "--port", $Port)
if ($Scenario) { $copyArguments += @("--scenario", $Scenario) }
if ($Run) { $copyArguments += "--run" }
& (Join-Path $PSScriptRoot "run.ps1") mock-workbook @copyArguments
if ($LASTEXITCODE -ne 0) {
    if ($Run) { Fail "J2D-SCR-053" "the unattended run of the scenario failed (see above and scripts\logs\manual-test.log)." }
    Fail "J2D-SCR-021" "the copy could not be prepared (see above)."
}
# The folder of a scenario is its own: its copy, its Rovo CSV, its results (the Python path writes its xlsx next to the copy).
$folder = Join-Path $root "out\manual_test"
$copyName = "jira2dashboard4manualtest.xlsm"
if ($Scenario) { $folder = Join-Path $folder $Scenario; $copyName = "jira2dashboard4$Scenario.xlsm" }
if ($Run) {
    $page = Join-Path $root ("out\manual_test\charts\" + ($Scenario -replace "-", "_") + "\index.html")
    Start-Process $page
    Write-Host "manual-test: the scenario $Scenario ran alone, nothing to click. The charts: $page"
    Write-Host "manual-test: the copy with its results: $(Join-Path $folder $copyName)"
    exit 0
}
if ($Scenario) { & (Join-Path $PSScriptRoot "start-mock.ps1") -Port $Port -Scenario $Scenario }
else { & (Join-Path $PSScriptRoot "start-mock.ps1") -Port $Port }
if ($LASTEXITCODE -ne 0) { Fail "J2D-SCR-022" "the fake Jira did not start (see above)." }

$copy = Join-Path $folder $copyName
if ($Python) {
    & (Join-Path $PSScriptRoot "run.ps1") extract --workbook $copy
    if ($LASTEXITCODE -ne 0) { Fail "J2D-SCR-023" "the Python extraction failed (see above and its log)." }
    $result = Join-Path $folder "jira2dashboard.xlsx"
    Start-Process excel.exe -ArgumentList "/x", "`"$result`""
    Write-Host "manual-test: $result opens in a new Excel: the data, the results and the charts, no macro."
    Write-Host "manual-test: stop the fake Jira with .\scripts\stop-mock.ps1 -Port $Port"
    exit 0
}
Write-Host "manual-test: Excel may crash right after the last search page (known defect). Save your other work first."
Start-Process excel.exe -ArgumentList "/x", "`"$copy`""
Write-Host "manual-test: $copy opens in a new Excel. Enable the content, then click the Dashboard buttons in this order:"
Write-Host "  1. Read the Jira configuration"
Write-Host "  2. Extract from Jira and compute"
Write-Host "manual-test: stop the fake Jira with .\scripts\stop-mock.ps1 -Port $Port"
exit 0
