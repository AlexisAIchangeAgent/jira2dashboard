<#
.SYNOPSIS
Starts the fake Jira (with the known data anomalies) in a window of its own and waits until it answers.

.DESCRIPTION
Runs ".\scripts\run.ps1 mock-jira --anomalies" in a minimized window (no window with -Hidden) and waits up to 60 s
until the port answers. It fails at once if another program already holds the port, or if the fake Jira stops at
once. The pid and the start time of the window are kept in %TEMP%\jira2dashboard\mock-jira-<port>.pid, so that
stop-mock.ps1 only ever kills what this script started. It prints the path and the MD5 and SHA256 hashes of the python.exe
that serves the fake Jira (the program an endpoint protection may have to allow); read only. Account of the fake Jira: mock@example.com / mock-token.
ASCII only: PowerShell 5.1 reads a file without BOM as ANSI.
USE CASES
  A fake Jira to click on            .\scripts\start-mock.ps1                       (then open the workbook: scripts\manual-test.ps1 does both)
  A small one, on another port       .\scripts\start-mock.ps1 -Port 8100 -Issues 150
  From a test or a script            .\scripts\start-mock.ps1 -Hidden               (no window; stop it with stop-mock.ps1)
  Port already held                  it says so and starts nothing: choose another -Port
  It stopped at once                 run .\scripts\run.ps1 mock-jira --port <port> to see the error
  Where is the log                   %LOCALAPPDATA%\jira2dashboard\logs\start-mock.log (last run only)


.PARAMETER Port
Port of the fake Jira, on 127.0.0.1. Default 8099.

.PARAMETER Issues
Number of generated issues, before the known anomalies are added. Default 1200.

.PARAMETER Hidden
No window at all (used by the tests). Without it the window is minimized and shows the address and the account.

.EXAMPLE
.\scripts\start-mock.ps1
Fake Jira on http://127.0.0.1:8099 with 1200 issues and the known anomalies.

.EXAMPLE
.\scripts\start-mock.ps1 -Port 8100 -Issues 150
A smaller fake Jira on another port. Stop it with .\scripts\stop-mock.ps1 -Port 8100.
.PARAMETER Help
Show this help. --help and -? do the same.
#>
[CmdletBinding(PositionalBinding = $false)]
param(
    [int]$Port = 8099,
    [int]$Issues = 1200,
    [switch]$Hidden,
    [switch]$Help,
    [Parameter(ValueFromRemainingArguments = $true)][string[]]$Rest = @()
)
if ($Help -or $Rest -contains "--help") { Get-Help $PSCommandPath -Detailed; exit 0 }
if ($Rest.Count -gt 0) { Write-Host "start-mock: unknown argument $($Rest -join ' '). See .\scripts\start-mock.ps1 --help"; exit 2 }
$ErrorActionPreference = "Stop"

# Log of this run (last run only): %LOCALAPPDATA%\jira2dashboard\logs\start-mock.log. A log failure never stops the script.
# Called by another script of this folder that already writes a log ($J2D_LogOwner, seen through the calling scope):
# no second transcript, the lines of this script go to the log of the caller.
if (-not (Get-Variable J2D_LogOwner -ErrorAction SilentlyContinue)) {
    try {
        $logDir = Join-Path $env:LOCALAPPDATA "jira2dashboard\logs"
        New-Item -ItemType Directory -Force $logDir | Out-Null
        Start-Transcript -Path (Join-Path $logDir "start-mock.log") -Force | Out-Null
        $J2D_LogOwner = "start-mock"
    } catch { }
}
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

function Show-Interpreter([int]$port) {
    # The program that serves the fake Jira, for an allow-list of an endpoint protection (the venv python.exe is only a
    # launcher: the listener is the real interpreter). Read only, and never an error.
    try {
        $listener = $null
        foreach ($entry in (netstat -ano -p tcp)) {
            if ($entry -match "^\s*TCP\s+127\.0\.0\.1:$port\s+\S+\s+LISTENING\s+(\d+)\s*$") { $listener = [int]$Matches[1]; break }
        }
        if (-not $listener) { throw "no listener" }
        $path = (Get-Process -Id $listener -ErrorAction Stop).Path
        if (-not $path) { throw "no path" }
        $md5 = (Get-FileHash $path -Algorithm MD5).Hash.ToLower()
        $sha = (Get-FileHash $path -Algorithm SHA256).Hash.ToLower()
        Write-Host "start-mock: python: $path (MD5 $md5, SHA256 $sha)"
    } catch {
        Write-Host "start-mock: python: not found (not an error)"
    }
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
Show-Interpreter $Port
Write-Host "start-mock: stop it with .\scripts\stop-mock.ps1 -Port $Port"
exit 0
