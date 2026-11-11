<#
.SYNOPSIS
Kills the fake Jira that start-mock.ps1 (or manual-test.ps1) started on a port.

.DESCRIPTION
Reads %TEMP%\jira2dashboard\mock-jira-<port>.pid, written by start-mock.ps1, and kills that process tree (its window,
the Python launcher and the server: taskkill /T) only if a process with that pid AND that start time still runs. A
program that reused the pid is left alone, and no other program on the port is ever looked at. Without a pid file it
says so and does nothing. ASCII only: PowerShell 5.1 reads a file without BOM as ANSI.
USE CASES
  After a manual try                 .\scripts\stop-mock.ps1
  A fake Jira on another port        .\scripts\stop-mock.ps1 -Port 8100
  Nothing to stop                    it says so and does nothing: a fake Jira that these scripts did not start is never touched
  Where is the log                   %LOCALAPPDATA%\jira2dashboard\logs\stop-mock.log (last run only)


.PARAMETER Port
Port of the fake Jira to stop. Default 8099.

.PARAMETER Help
Show this help. --help and -? do the same.

.EXAMPLE
.\scripts\stop-mock.ps1
Stops the fake Jira of port 8099.

.EXAMPLE
.\scripts\stop-mock.ps1 -Port 8100
Stops the fake Jira that start-mock.ps1 -Port 8100 started.
#>
[CmdletBinding(PositionalBinding = $false)]
param(
    [int]$Port = 8099,
    [switch]$Help,
    [Parameter(ValueFromRemainingArguments = $true)][string[]]$Rest = @()
)
if ($Help -or $Rest -contains "--help") { Get-Help $PSCommandPath -Detailed; exit 0 }
if ($Rest.Count -gt 0) { Write-Host "stop-mock: unknown argument $($Rest -join ' '). See .\scripts\stop-mock.ps1 --help"; exit 2 }
$ErrorActionPreference = "Stop"

# Log of this run (last run only): %LOCALAPPDATA%\jira2dashboard\logs\stop-mock.log. A log failure never stops the script.
try {
    $logDir = Join-Path $env:LOCALAPPDATA "jira2dashboard\logs"
    New-Item -ItemType Directory -Force $logDir | Out-Null
    Start-Transcript -Path (Join-Path $logDir "stop-mock.log") -Force | Out-Null
} catch { }
$pidFile = Join-Path (Join-Path $env:TEMP "jira2dashboard") "mock-jira-$Port.pid"

if (-not (Test-Path $pidFile)) {
    Write-Host "stop-mock: no fake Jira started by these scripts on port $Port."
    exit 0
}
$id, $start = (Get-Content $pidFile -Raw).Trim() -split " "
$process = Get-Process -Id $id -ErrorAction SilentlyContinue
if ($process -and $process.StartTime.ToFileTimeUtc() -eq [long]$start) {
    taskkill /PID $id /T /F | Out-Null
    Write-Host "stop-mock: fake Jira of port $Port stopped."
} else {
    Write-Host "stop-mock: the fake Jira of port $Port had already stopped."
}
Remove-Item $pidFile
exit 0
