# .\scripts\start-mock.ps1 [-Port 8099] [-Issues 1200] [-Hidden]
# Starts the fake Jira (scripts\run.ps1 mock-jira, with the known data anomalies) in a minimized window of its own
# (-Hidden: no window, for the tests) and waits until it answers. Stop it with scripts\stop-mock.ps1: the pid and the
# start time of that window are kept in %TEMP%\jira2dashboard\mock-jira-<port>.pid, so that the stop script only ever
# kills what this script started. ASCII only: PowerShell 5.1 reads a file without BOM as ANSI.
param(
    [int]$Port = 8099,
    [int]$Issues = 1200,
    [switch]$Hidden
)
$ErrorActionPreference = "Stop"
$pidFile = Join-Path (Join-Path $env:TEMP "jira2dashboard") "mock-jira-$Port.pid"

function Fail([string]$message) {
    Write-Host "start-mock: $message"
    exit 1
}

function Answers([int]$port) {
    $client = New-Object System.Net.Sockets.TcpClient
    try { return $client.ConnectAsync("127.0.0.1", $port).Wait(500) } catch { return $false } finally { $client.Close() }
}

function Started {
    # The window this script started for that port, if it still runs (same pid and same start time).
    if (-not (Test-Path $pidFile)) { return $null }
    $id, $start = (Get-Content $pidFile -Raw).Trim() -split " "
    $process = Get-Process -Id $id -ErrorAction SilentlyContinue
    if ($process -and $process.StartTime.ToFileTimeUtc() -eq [long]$start) { return $process }
    return $null
}

if (Started) {
    Write-Host "start-mock: a fake Jira started by these scripts already runs on http://127.0.0.1:$Port"
    exit 0
}
if (Answers $Port) { Fail "port $Port is held by another program: choose another one with -Port." }

$run = Join-Path $PSScriptRoot "run.ps1"
$style = if ($Hidden) { "Hidden" } else { "Minimized" }
$arguments = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$run`"", "mock-jira", "--port", $Port,
               "--issues", $Issues, "--anomalies")
$window = Start-Process powershell -ArgumentList $arguments -WindowStyle $style -PassThru
New-Item -ItemType Directory -Force (Split-Path -Parent $pidFile) | Out-Null
Set-Content -Path $pidFile -Value ("{0} {1}" -f $window.Id, $window.StartTime.ToFileTimeUtc()) -Encoding Ascii

$deadline = (Get-Date).AddSeconds(60)
while (-not (Answers $Port)) {
    if ($window.HasExited) {
        Remove-Item $pidFile
        Fail "the fake Jira stopped at once. Run .\scripts\run.ps1 mock-jira --port $Port to see why."
    }
    if ((Get-Date) -gt $deadline) {
        & (Join-Path $PSScriptRoot "stop-mock.ps1") -Port $Port
        Fail "no answer on port $Port after 60 s."
    }
    Start-Sleep -Milliseconds 300
}
Write-Host "start-mock: fake Jira on http://127.0.0.1:$Port ($Issues issues and the known anomalies)."
Write-Host "start-mock: stop it with .\scripts\stop-mock.ps1 -Port $Port"
exit 0
