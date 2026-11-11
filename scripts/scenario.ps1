<#
.SYNOPSIS
Plays the whole scenario against the fake Jira, from one command: fake Jira, filled copy of the workbook, extraction, checks, clean-up.

.DESCRIPTION
Runs, in this order, the scripts of this folder (the output of each one is in the log of this run):
  1. stop-mock.ps1           stops a fake Jira that these scripts started earlier on the port
  2. run.ps1 build           only when the workbook is missing (default out\release\jira2dashboard.xlsm)
  3. run.ps1 mock-workbook   the copy out\manual_test\jira2dashboard4manualtest.xlsm, Config, Iterations and PIs filled
                             for the fake Jira (the workbook given is never changed)
  4. start-mock.ps1 -Hidden  the fake Jira with the known data anomalies, no window
  5. run.ps1 extract         the Python path, no macro: Jira -> out\manual_test\jira2dashboard.xlsx (values and charts)
  6. check                   the xlsx exists and is not empty (the extraction already checks its own data)
  7. run.ps1 test-excel      only with -Excel: the scenarios that run the workbook macros against the fake Jira
  8. stop-mock.ps1           always at the end, also when a step failed (unless -KeepMock)
A step that fails stops the scenario: the next steps are skipped (the clean-up still runs). At the end a table lists each
step with its status and its time. The log of the run (last run only) is scripts\logs\scenario.log: its
lines "step N: name" and "scenario: ..." tell which step failed. The scripts it plays write into that same log (one transcript,
not one per script, with the output of Python): run_<command>.log, start-mock.log... keep the log of their last run on
their own.
Exit status: 0 when every step passed, 1 when a step failed, 2 for a wrong argument.
ASCII only: PowerShell 5.1 reads a file without BOM as ANSI.

USE CASES
  Check the chain by hand            .\scripts\scenario.ps1                     (fake Jira, Python extraction, clean-up)
  Check the macros too               .\scripts\scenario.ps1 -Excel              (adds the Excel scenarios: about 5 minutes)
  Look at the result in Excel        .\scripts\scenario.ps1 -KeepMock           (the fake Jira stays up; stop it with stop-mock.ps1)
  Try a downloaded workbook          .\scripts\scenario.ps1 -Workbook C:\Downloads\jira2dashboard.xlsm
  Port 8099 is already in use        .\scripts\scenario.ps1 -Port 8100
  A step failed                      read scripts\logs\scenario.log (the output of every step is in it)
  Click through the macros yourself  .\scripts\manual-test.ps1                  (interactive: opens the copy in Excel)

.PARAMETER Workbook
The workbook to try. Default out\release\jira2dashboard.xlsm (built if missing).

.PARAMETER Port
Port of the fake Jira, on 127.0.0.1. Default 8099.

.PARAMETER Excel
Also run the Excel scenarios (.\scripts\run.ps1 test-excel): they open the workbook in Excel and play the macros.

.PARAMETER KeepMock
Do not stop the fake Jira at the end: to look at it, or to play the macros by hand on the copy.

.PARAMETER Help
Show this help. --help and -? do the same.

.EXAMPLE
.\scripts\scenario.ps1
Plays the scenario with the Python extraction and prints the table of the steps.

.EXAMPLE
.\scripts\scenario.ps1 -Excel -Port 8100
The same on port 8100, then the Excel scenarios.
#>
[CmdletBinding(PositionalBinding = $false)]
param(
    [string]$Workbook = "",
    [int]$Port = 8099,
    [switch]$Excel,
    [switch]$KeepMock,
    [switch]$Help,
    [Parameter(ValueFromRemainingArguments = $true)][string[]]$Rest = @()
)
if ($Help -or $Rest -contains "--help") { Get-Help $PSCommandPath -Detailed; exit 0 }
if ($Rest.Count -gt 0) { Write-Host "scenario: J2D-SCR-001 - unknown argument $($Rest -join ' '). See .\scripts\scenario.ps1 --help"; exit 2 }
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

# Log of this run (last run only): scripts\logs\scenario.log. A log failure never stops the script.
# Called by another script of this folder that already writes a log ($J2D_LogOwner, seen through the calling scope):
# no second transcript, the lines of this script go to the log of the caller.
if (-not (Get-Variable J2D_LogOwner -ErrorAction SilentlyContinue)) {
    try {
        $logDir = Join-Path $PSScriptRoot "logs"
        New-Item -ItemType Directory -Force $logDir | Out-Null
        Start-Transcript -Path (Join-Path $logDir "scenario.log") -Force | Out-Null
        $J2D_LogOwner = "scenario"
    } catch { }
}

$results = New-Object System.Collections.ArrayList
$failed = $false
$number = 0

function Step([string]$name, [scriptblock]$action) {
    # One step: timed, logged ("step N: name" then "step N: ok|FAILED"), never run after a failed one.
    $script:number++
    if ($script:failed) {
        [void]$script:results.Add([pscustomobject]@{ Number = $script:number; Name = $name; Status = "skipped"; Seconds = 0 })
        Write-Host "scenario: step $($script:number): $name - skipped (an earlier step failed)"
        return
    }
    Write-Host "scenario: step $($script:number): $name"
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $global:LASTEXITCODE = 0
    $status = "ok"
    try {
        & $action
        if ($LASTEXITCODE -ne 0) { $status = "FAILED (exit code $LASTEXITCODE)" }
    } catch {
        $status = "FAILED ($($_.Exception.Message))"
    }
    $watch.Stop()
    if ($status -ne "ok") { $script:failed = $true }
    Write-Host "scenario: step $($script:number): $status"
    [void]$script:results.Add([pscustomobject]@{ Number = $script:number; Name = $name; Status = $status; Seconds = $watch.Elapsed.TotalSeconds })
}

$run = Join-Path $PSScriptRoot "run.ps1"
$copy = Join-Path $root "out\manual_test\jira2dashboard4manualtest.xlsm"
$result = Join-Path $root "out\manual_test\jira2dashboard.xlsx"
$source = $Workbook
if (-not $source) { $source = Join-Path $root "out\release\jira2dashboard.xlsm" }

Write-Host "scenario: workbook $source, fake Jira on port $Port$(if ($Excel) { ', with the Excel scenarios' })"
Step "stop an earlier fake Jira" { & (Join-Path $PSScriptRoot "stop-mock.ps1") -Port $Port }
Step "build the workbook if missing" {
    if (Test-Path $source) { Write-Host "scenario: $source exists, not built." }
    elseif ($Workbook) { throw "J2D-SCR-041 - the workbook $Workbook does not exist" }
    else { & $run build }
}
Step "fill a copy of the workbook for the fake Jira" { & $run mock-workbook --workbook $source --port $Port }
Step "start the fake Jira" { & (Join-Path $PSScriptRoot "start-mock.ps1") -Port $Port -Hidden }
Step "extract with Python (no macro)" { & $run extract --workbook $copy }
Step "check the xlsx" {
    if (-not (Test-Path $result)) { throw "J2D-SCR-042 - $result was not written" }
    if ((Get-Item $result).Length -eq 0) { throw "J2D-SCR-043 - $result is empty" }
    Write-Host ("scenario: {0} ({1} bytes)" -f $result, (Get-Item $result).Length)
}
if ($Excel) {
    Step "play the macros in Excel (test-excel)" { & $run test-excel }
}
if (-not $KeepMock) {
    $failedBefore = $failed
    $failed = $false                                   # the clean-up runs after a failure too
    Step "stop the fake Jira" { & (Join-Path $PSScriptRoot "stop-mock.ps1") -Port $Port }
    $failed = $failedBefore -or $failed
}

Write-Host ""
Write-Host "scenario: summary"
foreach ($r in $results) { Write-Host ("  {0,2}  {1,-46} {2,-34} {3,7:N1} s" -f $r.Number, $r.Name, $r.Status, $r.Seconds) }
if ($failed) {
    Write-Host "scenario: J2D-SCR-040 - FAILED. Log: $(Join-Path $PSScriptRoot 'logs\scenario.log')"
    exit 1
}
if ($KeepMock) { Write-Host "scenario: the fake Jira still runs on port ${Port}: stop it with .\scripts\stop-mock.ps1 -Port $Port" }
Write-Host "scenario: all steps passed."
exit 0
