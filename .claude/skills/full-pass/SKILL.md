---
name: full-pass
description: Run the full verification pass of jira2dashboard (coverage, vba-tests, test-excel, test-network, build) in order, tell the known Excel crash apart from a real failure, and report the result in French.
disable-model-invocation: true
---

# Full pass

The pass CLAUDE.md requires before finishing a change. Run the steps **in this order**, one at a time, from Bash, each
with its log in `$TEMP` (never in `out/`). Do not stop at the first red step: run them all, then report.

| # | Step | Command (from the repository folder) | Log |
|---|---|---|---|
| 1 | coverage | `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/run.ps1 coverage -q` | `$TEMP/fp_coverage.log` |
| 2 | vba-tests | `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/run.ps1 vba-tests` | `$TEMP/fp_vba.log` |
| 3 | test-excel | `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/run.ps1 test-excel -q` | `$TEMP/fp_excel.log` |
| 4 | test-network | `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/run.ps1 test-network -q` | `$TEMP/fp_network.log` |
| 5 | build | `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/run.ps1 build` | `$TEMP/fp_build.log` |

- Redirect outside PowerShell: `... > "$TEMP/fp_x.log" 2>&1; echo "exit $?"`. Never `cmd /c`, never `2>&1` inside
  PowerShell.
- Timeout 600000 ms per step. `test-excel` takes about 5 minutes: run it with `run_in_background` and wait for the
  notification; do not poll.
- Read each result from the log: the exit code, the pytest summary line (`passed`, `failed`, `error`), the coverage
  total (`Required test coverage of 90.0% reached`), `COMPILED_OK` and the marker coverage for `vba-tests`,
  `built ...` and `COMPILED_OK` for `build`.

## When an Excel step is red

1. List the failures: `grep -E "^(FAILED|ERROR)" <log>`. `Windows fatal exception: code 0x800706be/0x800706ba` is
   teardown noise, not a failure.
2. Read the end of the VBA log: `tail -5 out/inter/logs/jira2dashboard.log   # the folder logs next to the workbook under test`. If the last line is
   `END search page N: ...` (or the next run says `previous run did not end`), it is the **known open defect** (Excel
   crash after the last search page, handoff section 8b), not the change.
3. Rerun only the failing scenario once (`run.ps1 test-excel <file> -k "<name>" -q`).
4. Still red: before blaming the change, run the last commit the same number of times in a worktree, as CLAUDE.md says
   (`git worktree add --detach <scratchpad>/wt HEAD`, then pytest there with the main `.venv` python and
   `PYTHONPATH=<wt>/src`). Remove the worktree afterwards.
5. Never kill an `EXCEL.EXE` the tests did not start. Ask AL before reading the Windows event log.

## Report (in French, short sentences)

A table: step, result (vert / rouge / défaut connu), counts, log path. Then, for each red step, the failing tests and
the evidence (VBA log line, rerun result, HEAD comparison). Claim green only for steps whose log you read in this run.
If everything is green and the work is finished, commit and push as AL allows; never `release.ps1 -Publish`.
