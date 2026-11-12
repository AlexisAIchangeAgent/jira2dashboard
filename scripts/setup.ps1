<#
.SYNOPSIS
Creates .venv with Python 3.14 and installs the requirements.

.DESCRIPTION
Creates .venv (py -3.14 -m venv) if it does not exist, upgrades pip, then installs requirements.txt (pywin32: what the
workbook build and the Python extraction need, and pytest, pytest-bdd and pytest-cov, to run the tests). One file, one
install, on every machine. Checks at the end that pywin32 can be imported. PowerShell 5.1 compatible.
USE CASES
  Any machine                        .\scripts\setup.ps1                 (then .\scripts\run.ps1 test)
  Behind a proxy                     .\scripts\setup.ps1 -Proxy http://proxy.company.com:8080
  After a change of requirements     run it again: an existing .venv is kept, pip only installs what changed
  "py -3.14 failed"                  install Python 3.14 from python.org, then run it again
  Where is the log                   scripts\logs\setup.log (last run only)


.PARAMETER Proxy
Proxy for pip, for example http://proxy.company.com:8080. Empty: pip finds it itself.

.EXAMPLE
powershell -ExecutionPolicy Bypass -File scripts\setup.ps1
Everything, tests included.

.EXAMPLE
powershell -ExecutionPolicy Bypass -File scripts\setup.ps1 -Proxy http://proxy.company.com:8080
The same, behind a proxy.
.PARAMETER Help
Show this help. --help and -? do the same.
#>
[CmdletBinding(PositionalBinding = $false)]
param(
    [string]$Proxy = "",
    [switch]$Help,
    [Parameter(ValueFromRemainingArguments = $true)][string[]]$Rest = @()
)
if ($Help -or $Rest -contains "--help") { Get-Help $PSCommandPath -Detailed; exit 0 }
if ($Rest.Count -gt 0) { Write-Host "setup: J2D-SCR-001 - unknown argument $($Rest -join ' '). See .\scripts\setup.ps1 --help"; exit 2 }
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

# Log of this run (last run only): scripts\logs\setup.log. A log failure never stops the script.
# Called by another script of this folder that already writes a log ($J2D_LogOwner, seen through the calling scope):
# no second transcript, the lines of this script go to the log of the caller.
if (-not (Get-Variable J2D_LogOwner -ErrorAction SilentlyContinue)) {
    try {
        $logDir = Join-Path $PSScriptRoot "logs"
        New-Item -ItemType Directory -Force $logDir | Out-Null
        Start-Transcript -Path (Join-Path $logDir "setup.log") -Force | Out-Null
        $J2D_LogOwner = "setup"
    } catch { }
}

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    & py -3.14 -m venv .venv
    if ($LASTEXITCODE -ne 0) { throw "J2D-SCR-004 - py -3.14 -m venv failed (is Python 3.14 installed?)" }
}
$python = Join-Path $root ".venv\Scripts\python.exe"

$pipArgs = @()
if ($Proxy -ne "") { $pipArgs += @("--proxy", $Proxy) }

& $python -m pip install @pipArgs --upgrade pip
if ($LASTEXITCODE -ne 0) { throw "J2D-SCR-005 - pip upgrade failed" }

$req = "requirements.txt"
& $python -m pip install @pipArgs -r $req
if ($LASTEXITCODE -ne 0) { throw "J2D-SCR-006 - pip install -r $req failed" }

& $python -c "import win32com.client, pythoncom; print('pywin32 import OK')"
if ($LASTEXITCODE -ne 0) { throw "J2D-SCR-007 - pywin32 cannot be imported in .venv" }
Write-Host "Setup OK ($req)."
