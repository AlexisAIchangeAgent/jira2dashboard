---
name: read-logs
description: Read the logs of jira2dashboard after a run (test-excel, test-network, vba-tests, manual-test, scenario, extract, or a log AL brings back from the client) and report in French what failed with its support code, what was slow, what was retried and what did not end. Use when AL asks to look at the logs, why a run failed or stopped, or why it was slow.
---

# Read the logs

Every run leaves its logs on this machine; nothing leaves it. Read them, summarize, then answer AL in French, short
simple sentences, with the full path of each log you quote.

## Where the logs are

| Log | Path | Holds |
|---|---|---|
| VBA log of a workbook | `<folder of the .xlsm>\logs\jira2dashboard.log` (`out\inter\logs`, `out\manual_test\logs`, `out\release\logs`) | `RUN` (button, Excel, workbook, version), `PARAM` (inputs), `BEGIN`/`END ... \| 0.235 s`, `RETRY`, `ERROR <code>`, `RUN end ok\|stopped <code> \| s` |
| Python log | `<folder of the .xlsm>\logs\python.log` | the same for `run.ps1 extract`, every Config value but the token |
| Excel tests | `out\inter\logs\excel-test_<test>.log` | header `# test \| result \| time \| date`, its VBA lines, its `python.log` |
| VBA tests | `out\inter\logs\vba-tests.log` | `TEST <name>` / `TEST <name> done \| s`: after a crash the last `TEST` line names the test |
| Scripts | `scripts\logs\run_<command>.log`, `scripts\logs\<script>.log` | transcript with pytest's output and the `J2D-SCR-nnn` lines |

## Steps

1. Summary first, from the repository folder:
   `.venv\Scripts\python.exe .claude\skills\read-logs\summarize_logs.py` (default: the four log folders), or with the
   files or folders to read (a log from the client: its path).
2. For each failure or stop, read the lines around it in the log itself (the `BEGIN` without its `END` is the step that
   was running when Excel crashed or the run was stopped).
3. Give the meaning of each support code: `.venv\Scripts\python.exe -m jira2dashboard.error_codes` (cause, check, what to ask).
4. Report: what ran (version of the workbook from the `RUN` line), what failed and why, the slowest steps, the retries,
   a run that did not end (`previous run did not end`), and what to do next.

## Rules

- `python.log` holds the JQL and Config values in clear (e-mail, address of Jira): quote the lines needed, never the whole
  file, and never paste a log into anything outside this machine.
- The `Windows fatal exception: code 0x800706be` lines after Excel tests are teardown noise: only the test result counts.
- An Excel crash (`RPC server unavailable`, `0x800706BE`) in `test_build_workbook` or right after the last search page
  is the known defect of CLAUDE.md: say so, do not chase it here.
