# Create .venv with Python 3.14 and install the requirements. PowerShell 5.1 compatible.
#   powershell -ExecutionPolicy Bypass -File scripts\setup.ps1 [-Dev] [-Proxy http://proxy:8080]
param(
    [switch]$Dev,
    [string]$Proxy = ""
)
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
