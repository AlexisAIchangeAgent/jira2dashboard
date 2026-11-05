<#
.SYNOPSIS
Activates .venv in the current PowerShell session (optional: scripts\run.ps1 needs no activation).

.DESCRIPTION
Dot-source it: . .\scripts\activate.ps1
It runs .venv\Scripts\Activate.ps1, sets PYTHONIOENCODING=utf-8, switches the console to UTF-8 (chcp 65001) and sets
PYTHONPATH to src, so that "python -m jira2dashboard..." works by hand. Create .venv first with scripts\setup.ps1.
PowerShell 5.1 compatible.

.PARAMETER Help
Show this help. --help and -? do the same.

.EXAMPLE
. .\scripts\activate.ps1
Then: python -m jira2dashboard.mockjira --help
#>
# Dot-sourced: never "exit" here (it would close the session), "return" instead. No param block: $args holds --help.
if (@("--help", "-h", "-help", "/?") -contains "$($args | Select-Object -First 1)") {
    Get-Help $PSCommandPath -Detailed
    return
}
$root = Split-Path -Parent $PSScriptRoot
$activate = Join-Path $root ".venv\Scripts\Activate.ps1"
if (-not (Test-Path $activate)) {
    Write-Error "No .venv yet: run scripts\setup.ps1 first."
} else {
    . $activate
    $env:PYTHONIOENCODING = "utf-8"
    chcp 65001 | Out-Null
    $env:PYTHONPATH = (Join-Path $root "src")
    Write-Host "jira2dashboard venv active (PYTHONPATH=$env:PYTHONPATH)."
}
