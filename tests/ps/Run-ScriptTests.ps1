<#
.SYNOPSIS
Runs the Pester tests of scripts\*.ps1 and measures how much of each script they execute.

.DESCRIPTION
Copies scripts\*.ps1 into a sandbox folder, runs tests\ps\Scripts.Tests.ps1 (Pester 3.4, in process, with fake git, gh, py and
python: no .venv, no Excel, no network, nothing of this repository is touched) and prints the coverage of the commands of
the scripts. It fails when a test fails or when the coverage is under the minimum (93 % by default). The launchers
scripts\*.cmd are tested apart (tests\unit\test_cmd_launchers.py, run by .\scripts\run.ps1 test).
Run it with: .\scripts\run.ps1 test-scripts

USE CASES
  After a change of a script         .\scripts\run.ps1 test-scripts   (about a minute; the full pass runs it too)
  See what is not covered            .\scripts\run.ps1 test-scripts -Verbose   (every command no test executes)
  A stricter bar                     .\scripts\run.ps1 test-scripts -Minimum 97

.PARAMETER Minimum
The coverage required, in percent of the commands of scripts\*.ps1. Default 93.

.PARAMETER Filter
Only the Describe blocks whose name matches this wildcard (for example "setup.ps1" or "*mock*"). The coverage is then partial: the minimum is not checked.

.PARAMETER Help
Show this help. --help and -? do the same.

.EXAMPLE
.\scripts\run.ps1 test-scripts
Runs the tests and the coverage; exit code 1 under 93 %.
#>
[CmdletBinding(PositionalBinding = $false)]
param(
    [int]$Minimum = 93,
    [string]$Filter = "",
    [switch]$Help,
    [Parameter(ValueFromRemainingArguments = $true)][string[]]$Rest = @()
)
if ($Help -or $Rest -contains "--help") { Get-Help $PSCommandPath -Detailed; exit 0 }
if ($Rest.Count -gt 0) { Write-Host "test-scripts: unknown argument $($Rest -join ' ')"; exit 2 }
$ErrorActionPreference = "Stop"
$repo = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
Import-Module Pester -ErrorAction Stop

# Under out\inter, not in %TEMP%: %TEMP% may be an 8.3 short path (C:\Users\ALEXIS~1), while a script sees its own folder under
# the long name, and the fake python.exe (a function named by its full path) must match the path the script computes.
$base = Join-Path $repo ("out\inter\ps_" + [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Force (Join-Path $base "scripts") | Out-Null
Copy-Item (Join-Path $repo "scripts\*.ps1") (Join-Path $base "scripts")
$env:J2D_PS_BASE = $base
try {
    $files = @(Get-ChildItem (Join-Path $base "scripts\*.ps1") | ForEach-Object { $_.FullName })
    $pester = @{ Path = (Join-Path $repo "tests\ps"); CodeCoverage = $files; PassThru = $true }
    if ($Filter) { $pester.TestName = $Filter }
    $result = Invoke-Pester @pester
    $coverage = $result.CodeCoverage
    $analyzed = $coverage.NumberOfCommandsAnalyzed
    $executed = $coverage.NumberOfCommandsExecuted
    $percent = if ($analyzed) { [math]::Round(100.0 * $executed / $analyzed, 1) } else { 0 }
    Write-Host ""
    Write-Host "Coverage of scripts\*.ps1 by file:"
    foreach ($file in $files) {
        $name = Split-Path -Leaf $file
        $all = @($coverage.HitCommands | Where-Object { $_.File -eq $file }).Count + @($coverage.MissedCommands | Where-Object { $_.File -eq $file }).Count
        $hit = @($coverage.HitCommands | Where-Object { $_.File -eq $file }).Count
        $filePercent = if ($all) { [math]::Round(100.0 * $hit / $all, 1) } else { 100 }
        Write-Host ("  {0,-18} {1,5} %  ({2} of {3} commands)" -f $name, $filePercent, $hit, $all)
    }
    Write-Host ("Total: {0} % ({1} of {2} commands), minimum {3} %." -f $percent, $executed, $analyzed, $Minimum)
    if ($VerbosePreference -eq "Continue") {
        foreach ($missed in $coverage.MissedCommands) {
            Write-Host ("  not executed: {0}:{1}  {2}" -f (Split-Path -Leaf $missed.File), $missed.Line, $missed.Command)
        }
    }
    if ($result.FailedCount -gt 0) { Write-Host "test-scripts: $($result.FailedCount) test(s) failed."; exit 1 }
    if (-not $Filter -and $percent -lt $Minimum) { Write-Host "test-scripts: coverage $percent % is under $Minimum %."; exit 1 }
    Write-Host "test-scripts: $($result.PassedCount) tests passed."
    exit 0
} finally {
    Remove-Item $base -Recurse -Force -ErrorAction SilentlyContinue
}
