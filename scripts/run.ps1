<#
.SYNOPSIS
Runs a command of jira2dashboard with the Python of .venv (tests, build of the workbook, fake Jira...).

.DESCRIPTION
.\scripts\run.ps1 <command> [arguments of the command]

No venv activation is needed: the script uses .venv\Scripts\python.exe and sets PYTHONPATH to src. The arguments
after the command reach Python (or pytest) untouched, so ".\scripts\run.ps1 <command> --help" shows the options of
that command. Create .venv first with scripts\setup.ps1 (-Dev for the test tools).

Commands:
  test           fast unit and Gherkin tests, no Excel (pytest arguments allowed)
  coverage       the same, plus the coverage of the pure code; fails under 93 %
  test-excel     the scenarios that drive Excel (marker excel, no network)
  test-network   the Excel scenarios that need the network (marker network)
  vba-tests      build, compile and run the VBA test suite (--offline, --bisect, --out name.xlsm)
  check-machine  22 checks of this machine: Excel, access to the VBA project, compile, macro of the workbook...
  diagnose       a report to paste into a ticket when the macros are blocked (hashes, origin, Office, security products)
  mock-jira      the fake Jira on http://127.0.0.1:8099 (--anomalies, --issues, --port...)
  mock-workbook  a copy of a workbook with Config, Iterations and PIs filled for the fake Jira
  build          the workbook to deliver: out\release\jira2dashboard.xlsm (--out name.xlsm: in out\inter)
  clean          empty out\inter, except the test caches (build and vba-tests do it first)
  extract        without VBA: read Jira with Python, write data, results and charts to jira2dashboard.xlsx
  test-scripts   Pester tests of scripts\*.ps1 and their coverage; fails under 93 % (-Filter, -Minimum)

USE CASES
  First time on a machine            powershell -ExecutionPolicy Bypass -File scripts\setup.ps1 -Dev
  While you code                     .\scripts\run.ps1 test            (fast, no Excel); one test: .\scripts\run.ps1 test tests/unit/test_x.py::test_name
  Before a commit                    .\scripts\run.ps1 coverage ; vba-tests ; test-excel ; test-network ; test-scripts ; build
  Check the scripts of this folder   .\scripts\run.ps1 test-scripts        (Pester, coverage of the .ps1 under 93 % fails)
  The workbook to deliver            .\scripts\run.ps1 build           then .\scripts\release.ps1 -Version v1.0.0
  Try the macros against a fake Jira .\scripts\manual-test.ps1  (by hand)   or   .\scripts\scenario.ps1  (unattended)
  Extract without Excel macros       .\scripts\run.ps1 extract --workbook C:\Data\jira2dashboard.xlsm
  The macros are blocked at a client .\scripts\run.ps1 check-machine ; .\scripts\run.ps1 diagnose   (a report to paste into the ticket)
  Where are the logs                 %LOCALAPPDATA%\jira2dashboard\logs\<script>.log (last run only); the workbook logs in its own logs folder

.PARAMETER Help
Show this help. --help, -h and -? do the same.

.EXAMPLE
.\scripts\run.ps1 test
Runs the fast tests.

.EXAMPLE
.\scripts\run.ps1 build --help
Shows the options of the command build.

.EXAMPLE
.\scripts\run.ps1 test-excel tests/steps/test_flow_excel.py -k "contr" -q
Runs some Excel scenarios, with pytest arguments.
#>
# Plain $args (no param block) so that "--out", "--offline"... reach Python untouched.
$commands = [ordered]@{
    "test"          = "fast unit and Gherkin tests, no Excel (pytest arguments allowed)"
    "coverage"      = "the same, plus the coverage of the pure code; fails under 93 %"
    "test-excel"    = "the scenarios that drive Excel (marker excel, no network)"
    "test-network"  = "the Excel scenarios that need the network (marker network)"
    "vba-tests"     = "build, compile and run the VBA test suite (--offline, --bisect, --out name.xlsm)"
    "check-machine" = "22 checks of this machine: Excel, access to the VBA project, compile, macro of the workbook..."
    "diagnose"      = "a report to paste into a ticket when the macros are blocked (hashes, origin, Office, security products)"
    "mock-jira"     = "the fake Jira on http://127.0.0.1:8099 (--anomalies, --issues, --port...)"
    "mock-workbook" = "a copy of a workbook with Config, Iterations and PIs filled for the fake Jira"
    "build"         = "the workbook to deliver: out\release\jira2dashboard.xlsm (--out name.xlsm: in out\inter)"
    "clean"         = "empty out\inter, except the test caches (build and vba-tests do it first)"
    "extract"       = "without VBA: read Jira with Python, write data, results and charts to jira2dashboard.xlsx"
    "test-scripts"  = "Pester tests of scripts\*.ps1 and their coverage; fails under 93 % (-Filter, -Minimum)"
}
if ($args.Count -lt 1 -or @("--help", "-h", "-help", "help", "/?") -contains $args[0]) {
    Write-Host "Usage: .\scripts\run.ps1 <command> [arguments]   (.\scripts\run.ps1 <command> --help: its options)"
    Write-Host ""
    foreach ($name in $commands.Keys) { Write-Host ("  {0,-14} {1}" -f $name, $commands[$name]) }
    Write-Host ""
    Write-Host "Use cases:"
    Write-Host "  while you code          .\scripts\run.ps1 test"
    Write-Host "  before a commit         coverage, vba-tests, test-excel, test-network, test-scripts, build"
    Write-Host "  try the macros          .\scripts\manual-test.ps1 (by hand) or .\scripts\scenario.ps1 (unattended)"
    Write-Host "  a client machine        check-machine, then diagnose"
    if ($args.Count -lt 1) { exit 2 }
    exit 0
}
if (-not $commands.Contains($args[0])) {
    Write-Host "run: J2D-SCR-002 - unknown command '$($args[0])'. See .\scripts\run.ps1 --help"
    exit 2
}
$command = $args[0]
$rest = @()
if ($args.Count -gt 1) { $rest = $args[1..($args.Count - 1)] }

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$python = Join-Path $root ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) { throw "J2D-SCR-003 - No .venv: run scripts\setup.ps1 -Dev first." }
Set-Location $root
# Log of this run (last run only): %LOCALAPPDATA%\jira2dashboard\logs\run.log. A log failure never stops the script.
# Called by another script of this folder that already writes a log ($J2D_LogOwner, seen through the calling scope):
# no second transcript, the lines of this script go to the log of the caller.
if (-not (Get-Variable J2D_LogOwner -ErrorAction SilentlyContinue)) {
    try {
        $logDir = Join-Path $env:LOCALAPPDATA "jira2dashboard\logs"
        New-Item -ItemType Directory -Force $logDir | Out-Null
        Start-Transcript -Path (Join-Path $logDir "run.log") -Force | Out-Null
        $J2D_LogOwner = "run"
    } catch { }
}
$env:PYTHONIOENCODING = "utf-8"
$env:PYTHONPATH = (Join-Path $root "src")

switch ($command) {
    "test"          { & $python -m pytest @rest }
    "coverage"      { & $python -m pytest --cov --cov-report=term-missing @rest }
    "test-excel"    { & $python -m pytest -m "excel and not network" @rest }
    "test-network"  { & $python -m pytest -m "excel and network" @rest }
    "vba-tests"     { & $python -m jira2dashboard.testbench.run_vba_tests @rest }
    "check-machine" { & $python -m jira2dashboard.cli.check_machine @rest }
    "diagnose"      { & $python -m jira2dashboard.cli.diagnose @rest }
    "mock-jira"     { & $python -m jira2dashboard.mockjira @rest }
    "mock-workbook" { & $python -m jira2dashboard.mock_workbook @rest }
    "build"         { & $python -m jira2dashboard.build_workbook @rest }
    "clean"         { & $python -m jira2dashboard.clean @rest }
    "extract"       { & $python -m jira2dashboard.standalone @rest }
    "test-scripts"  { & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "tests\ps\Run-ScriptTests.ps1") @rest }
}
exit $LASTEXITCODE
