<#
.SYNOPSIS
Installs, checks, tests and builds everything from one command: .venv, the machine, the tests, the workbook (.xlsm) and the results file (.xlsx).

.DESCRIPTION
Runs, in this order, the scripts of this folder (the output of each one is in the log of this run):
  1. setup.ps1               .venv with Python 3.14, then requirements.txt (pywin32 and the test tools)
  2. run.ps1 check-machine   the checks of the machine: Excel, access to the VBA project, compile, macro of a workbook.
                             It changes nothing on the machine: when a check fails, the message says what to set
  3. run.ps1 test            the fast tests, no Excel (skipped with -SkipTests)
  4. run.ps1 vba-tests       only with -Full: the VBA tests, in Excel (about 5 minutes)
  5. run.ps1 test-excel      only with -Full: the scenarios that play the macros in Excel (about 5 minutes)
  6. run.ps1 build           the workbook to deliver, out\release\jira2dashboard.xlsm
  7. check                   the .xlsm exists and is not empty (an endpoint protection can remove a new file with macros)
  8. scenario.ps1            the fake Jira, a filled copy of the workbook and the Python extraction: it writes the
                             .xlsx, out\manual_test\jira2dashboard.xlsx, proof that the whole chain works on this machine
A step that fails stops the run: the next steps are skipped. At the end a table lists each step with its status and its
time, then the files that were made with their size (and the MD5 of the .xlsm, for an allow-list). The log of the run (last run
only) is scripts\logs\bootstrap.log: its lines "bootstrap: step N: name" tell which step failed. The scripts it runs write
into that same log (one transcript, with the output of Python); run_<command>.log keeps the log of their last run on their own.
The .xlsx of the real Jira is made later, on a workbook that you filled: .\scripts\run.ps1 extract --workbook <the .xlsm>.
Exit status: 0 when every step passed, 1 when a step failed, 2 for a wrong argument.
ASCII only: PowerShell 5.1 reads a file without BOM as ANSI.

USE CASES
  A new machine                      .\scripts\bootstrap.ps1                    (install, check, fast tests, build, .xlsx)
  Behind a proxy                     .\scripts\bootstrap.ps1 -Proxy http://proxy.company.com:8080
  Everything, the Excel tests too    .\scripts\bootstrap.ps1 -Full              (about 15 minutes)
  Install and build only             .\scripts\bootstrap.ps1 -SkipTests
  Port 8099 is already in use        .\scripts\bootstrap.ps1 -Port 8100
  A step failed                      read scripts\logs\bootstrap.log, then run it again: every step can be repeated
  Double-click                       scripts\bootstrap.cmd

.PARAMETER Proxy
Proxy for pip, for example http://proxy.company.com:8080. Empty: pip finds it itself.

.PARAMETER Port
Port of the fake Jira that makes the .xlsx, on 127.0.0.1. Default 8099.

.PARAMETER Full
Also run the VBA tests and the Excel scenarios (run.ps1 vba-tests and test-excel).

.PARAMETER SkipTests
Run no test at all (not even the fast ones): install, check the machine, build.

.PARAMETER Help
Show this help. --help and -? do the same.

.EXAMPLE
.\scripts\bootstrap.ps1
Installs, checks, tests, builds the .xlsm and the .xlsx, and prints the table of the steps.

.EXAMPLE
.\scripts\bootstrap.ps1 -Full -Proxy http://proxy.company.com:8080
The same behind a proxy, with the VBA tests and the Excel scenarios.
#>
[CmdletBinding(PositionalBinding = $false)]
param(
    [string]$Proxy = "",
    [int]$Port = 8099,
    [switch]$Full,
    [switch]$SkipTests,
    [switch]$Help,
    [Parameter(ValueFromRemainingArguments = $true)][string[]]$Rest = @()
)
if ($Help -or $Rest -contains "--help") { Get-Help $PSCommandPath -Detailed; exit 0 }
if ($Rest.Count -gt 0) { Write-Host "bootstrap: J2D-SCR-001 - unknown argument $($Rest -join ' '). See .\scripts\bootstrap.ps1 --help"; exit 2 }
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

# Log of this run (last run only): scripts\logs\bootstrap.log. A log failure never stops the script.
# Called by another script of this folder that already writes a log ($J2D_LogOwner, seen through the calling scope):
# no second transcript, the lines of this script go to the log of the caller.
if (-not (Get-Variable J2D_LogOwner -ErrorAction SilentlyContinue)) {
    try {
        $logDir = Join-Path $PSScriptRoot "logs"
        New-Item -ItemType Directory -Force $logDir | Out-Null
        Start-Transcript -Path (Join-Path $logDir "bootstrap.log") -Force | Out-Null
        $J2D_LogOwner = "bootstrap"
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
        Write-Host "bootstrap: step $($script:number): $name - skipped (an earlier step failed)"
        return
    }
    Write-Host "bootstrap: step $($script:number): $name"
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
    Write-Host "bootstrap: step $($script:number): $status"
    [void]$script:results.Add([pscustomobject]@{ Number = $script:number; Name = $name; Status = $status; Seconds = $watch.Elapsed.TotalSeconds })
}

$run = Join-Path $PSScriptRoot "run.ps1"
$xlsm = Join-Path $root "out\release\jira2dashboard.xlsm"
$xlsx = Join-Path $root "out\manual_test\jira2dashboard.xlsx"

Write-Host "bootstrap: root $root, fake Jira on port $Port$(if ($Full) { ', with the VBA and Excel tests' })$(if ($SkipTests) { ', no test' })"
Step "install: .venv and requirements.txt" {
    if ($Proxy -ne "") { & (Join-Path $PSScriptRoot "setup.ps1") -Proxy $Proxy } else { & (Join-Path $PSScriptRoot "setup.ps1") }
}
Step "check the machine" { & $run check-machine }
if (-not $SkipTests) {
    Step "fast tests (no Excel)" { & $run test }
    if ($Full) {
        Step "VBA tests (Excel)" { & $run vba-tests }
        Step "Excel scenarios (test-excel)" { & $run test-excel }
    }
}
Step "build the workbook (.xlsm)" { & $run build }
Step "check the .xlsm" {
    if (-not (Test-Path $xlsm)) { throw "J2D-SCR-051 - $xlsm is missing after the build: an endpoint protection probably removed it (J2D-SYS-012)" }
    if ((Get-Item $xlsm).Length -eq 0) { throw "J2D-SCR-051 - $xlsm is empty" }
    Write-Host ("bootstrap: {0} ({1} bytes)" -f $xlsm, (Get-Item $xlsm).Length)
}
Step "build the .xlsx (fake Jira, Python extraction)" { & (Join-Path $PSScriptRoot "scenario.ps1") -Port $Port }

Write-Host ""
Write-Host "bootstrap: summary"
foreach ($r in $results) { Write-Host ("  {0,2}  {1,-46} {2,-34} {3,7:N1} s" -f $r.Number, $r.Name, $r.Status, $r.Seconds) }
if ($failed) {
    Write-Host "bootstrap: J2D-SCR-050 - FAILED. Log: $(Join-Path $PSScriptRoot 'logs\bootstrap.log')"
    exit 1
}
Write-Host "bootstrap: files made"
foreach ($file in $xlsm, $xlsx) {
    if (Test-Path $file) { Write-Host ("  {0} ({1} bytes)" -f $file, (Get-Item $file).Length) }
}
if (Test-Path $xlsm) { Write-Host ("  MD5 of jira2dashboard.xlsm: {0}" -f (Get-FileHash $xlsm -Algorithm MD5).Hash.ToLower()) }
Write-Host "bootstrap: all steps passed."
exit 0
