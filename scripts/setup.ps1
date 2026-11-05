<#
.SYNOPSIS
Creates .venv with Python 3.14 and installs the requirements.

.DESCRIPTION
Creates .venv (py -3.14 -m venv) if it does not exist, upgrades pip, then installs requirements.txt (pywin32: what the
workbook build and the Python extraction need) or, with -Dev, requirements-dev.txt (plus pytest, pytest-bdd and
pytest-cov, to run the tests). Checks at the end that pywin32 can be imported. PowerShell 5.1 compatible.

.PARAMETER Dev
Also install the test tools (requirements-dev.txt).

.PARAMETER Proxy
Proxy for pip, for example http://proxy.company.com:8080. Empty: pip finds it itself.

.EXAMPLE
powershell -ExecutionPolicy Bypass -File scripts\setup.ps1 -Dev
The developer machine: everything, tests included.

.EXAMPLE
powershell -ExecutionPolicy Bypass -File scripts\setup.ps1 -Proxy http://proxy.company.com:8080
A client machine behind a proxy: what the build and the extraction need.
.PARAMETER Help
Show this help. --help and -? do the same.
#>
[CmdletBinding(PositionalBinding = $false)]
param(
    [switch]$Dev,
    [string]$Proxy = "",
    [switch]$Help,
    [Parameter(ValueFromRemainingArguments = $true)][string[]]$Rest = @()
)
if ($Help -or $Rest -contains "--help") { Get-Help $PSCommandPath -Detailed; exit 0 }
if ($Rest.Count -gt 0) { Write-Host "setup: unknown argument $($Rest -join ' '). See .\scripts\setup.ps1 --help"; exit 2 }
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    & py -3.14 -m venv .venv
    if ($LASTEXITCODE -ne 0) { throw "py -3.14 -m venv failed (is Python 3.14 installed?)" }
}
$python = Join-Path $root ".venv\Scripts\python.exe"

$pipArgs = @()
if ($Proxy -ne "") { $pipArgs += @("--proxy", $Proxy) }

& $python -m pip install @pipArgs --upgrade pip
if ($LASTEXITCODE -ne 0) { throw "pip upgrade failed" }

$req = "requirements.txt"
if ($Dev) { $req = "requirements-dev.txt" }
& $python -m pip install @pipArgs -r $req
if ($LASTEXITCODE -ne 0) { throw "pip install -r $req failed" }

& $python -c "import win32com.client, pythoncom; print('pywin32 import OK')"
if ($LASTEXITCODE -ne 0) { throw "pywin32 cannot be imported in .venv" }
Write-Host "Setup OK ($req)."
