---
name: excel-hygiene
description: Keep the Excel processes of jira2dashboard clean around any run that starts Excel (vba-tests, test-excel, test-network, build, release dry run, mock-workbook, scenario). Count EXCEL.EXE before and after, tell the Excel the run started from an older one by its start time, stop only that one by PID. Use before and after every such run, and when a run crashed or hung.
user-invocable: false
---

# Excel hygiene (AL's rule)

After every Excel run, even a green one, `tasklist | grep -c EXCEL` must be back to what it was before. Excel is
not always AL's to kill: the user may have their own open (the one `manual-test.ps1` opens is AL's, leave it).

## Before the run

1. `tasklist | grep -c EXCEL` and note the number. Also note the time (`date "+%H:%M:%S"`).
2. If it is not 0, list who they are so you can tell them apart later:
   `powershell -NoProfile -Command "Get-Process EXCEL | Select Id,StartTime,MainWindowTitle | Format-List"`
   (the PowerShell tool is slow here: Bash `tasklist` for the count, this one-liner only for the details).

## After the run

1. `tasklist | grep -c EXCEL` again. Same number: done, say nothing.
2. More: a leftover. List it as above and decide by its **start time and title**:

| Title | Start time | It is | Action |
|---|---|---|---|
| `Excel` or `... - Récupération` | after the start of the run (often the same second as a crash) | the recovery Excel Office started after a crash of the run | stop it by PID |
| `Microsoft Visual Basic` | after the start of the run | a modal VBA error dialog of the run (untrapped error): the run is hung | stop it by PID (AL allows it) |
| anything | before the run | not yours (AL's, or an older one) | **leave it**, tell AL if it looks hung |

3. Stop by PID with PowerShell, never by name and never `taskkill /PID` from Git Bash (the `/PID` switch is
   rewritten into a path; with `MSYS_NO_PATHCONV=1` it works, but PowerShell is simpler):
   `Stop-Process -Id <pid>`. The hook `excel_kill_guard.py` asks before `taskkill /IM` and `Stop-Process -Name`: that is
   on purpose, do not look for a way around it.
4. Check the count again. Often the recovery Excel is gone by itself a few seconds later: look again before stopping it
   (`Stop-Process` then says the process cannot be found: nothing to do).

## Hung run

A run that does not move for long (`vba-tests` with no new line in `out\inter\logs\vba-tests.log`, `test-excel` without a
new `excel-test_*.log`) is usually an Excel on a VBA error dialog. Run `vba-tests` with `timeout 420` so the shell
does not wait forever. Find the Excel whose title is `Microsoft Visual Basic` and whose start time matches the run,
stop that one by PID, then read the end of the VBA log (`read-logs`) for the line before the stop.

## Report

One line in French when something was stopped: the PID, its title, its start time, why it was the run's. Nothing when the
count is back to the start. Never claim "Excel clean" without having read the count after the run.
