# .\scripts\run.ps1 test | coverage | test-excel | test-network | vba-tests | check-machine | mock-jira | build | clean  [extra args]
# No venv activation needed. Plain $args (no [Parameter]) so that "--out", "--offline"... reach Python untouched.
$commands = @("test", "coverage", "test-excel", "test-network", "vba-tests", "check-machine", "mock-jira", "build", "clean")
if ($args.Count -lt 1 -or $commands -notcontains $args[0]) {
    Write-Host ("Usage: .\scripts\run.ps1 <" + ($commands -join " | ") + "> [extra args]")
    exit 2
}
$command = $args[0]
$rest = @()
if ($args.Count -gt 1) { $rest = $args[1..($args.Count - 1)] }

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$python = Join-Path $root ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) { throw "No .venv: run scripts\setup.ps1 -Dev first." }
Set-Location $root
$env:PYTHONIOENCODING = "utf-8"
$env:PYTHONPATH = (Join-Path $root "src")

switch ($command) {
    "test"          { & $python -m pytest @rest }
    "coverage"      { & $python -m pytest --cov --cov-report=term-missing @rest }
    "test-excel"    { & $python -m pytest -m "excel and not network" @rest }
    "test-network"  { & $python -m pytest -m "excel and network" @rest }
    "vba-tests"     { & $python -m jira2dashboard.testbench.run_vba_tests @rest }
    "check-machine" { & $python -m jira2dashboard.cli.check_machine @rest }
    "mock-jira"     { & $python -m jira2dashboard.mockjira @rest }
    "build"         { & $python -m jira2dashboard.build_workbook @rest }
    "clean"         { & $python -m jira2dashboard.clean @rest }
}
exit $LASTEXITCODE
