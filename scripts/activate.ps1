# Dot-source it:  . .\scripts\activate.ps1     (PowerShell 5.1 compatible)
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
