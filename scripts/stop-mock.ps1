# .\scripts\stop-mock.ps1 [-Port 8099]
# Kills the fake Jira that scripts\start-mock.ps1 (or scripts\manual-test.ps1) started on that port: its window, the
# Python launcher and the server (taskkill /T). It reads %TEMP%\jira2dashboard\mock-jira-<port>.pid and kills only a
# process with that pid AND that start time, so a program that reused the pid is left alone. It never looks for other
# programs on the port. ASCII only: PowerShell 5.1 reads a file without BOM as ANSI.
param(
    [int]$Port = 8099
)
$ErrorActionPreference = "Stop"
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
