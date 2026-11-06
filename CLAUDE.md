# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

**jira2dashboard** builds an Excel workbook (`.xlsm`) that extracts Jira data, computes the six SAFe flow metrics
(Velocity, Flow Time, Flow Load, Flow Efficiency, Flow Distribution, Flow Predictability) and shows them on a dashboard
with native charts. Python (pywin32 + stdlib at runtime) **builds** the workbook and its VBA; the delivered workbook then
works **without Python**: VBA reads Jira, builds the tables, formulas and charts.

Platform: Windows 11, PowerShell 5.1, Python 3.14, Excel Microsoft 365 (French UI on the dev machine). Runtime = pywin32 +
stdlib only; pytest, pytest-bdd and pytest-cov are dev extras. The repo is self-contained: child Excel processes import
`jira2dashboard` through `PYTHONPATH=<repo>\src` (no `pip install`).

Docs and code are in English. Conversation with the user (AL) is in French, short simple sentences.

## Commands

Always go through `scripts\run.ps1` (PowerShell 5.1, no venv activation needed); extra arguments reach pytest/Python.

```powershell
powershell -ExecutionPolicy Bypass -File scripts\setup.ps1 -Dev   # create .venv (py -3.14) and install
.\scripts\run.ps1 test            # fast unit + Gherkin tests, no Excel (the inner loop)
.\scripts\run.ps1 coverage        # same + coverage of the pure code; fails under 90 %
.\scripts\run.ps1 test-excel      # scenarios that drive Excel (marker `excel`, no network)
.\scripts\run.ps1 test-network    # Excel + network scenarios (marker `network`)
.\scripts\run.ps1 vba-tests       # build + compile + run the VBA test suite; --offline, --bisect, --out name.xlsm
.\scripts\run.ps1 check-machine   # 21 checks of the machine (Excel, VBA trust...)
.\scripts\run.ps1 mock-jira       # fake Jira on http://127.0.0.1:8099 (--anomalies: with the known data anomalies)
.\scripts\run.ps1 build           # delivery workbook out\release\jira2dashboard.xlsm (--out name.xlsm: to out\inter)
.\scripts\run.ps1 clean           # empty out\inter except the test caches (build and vba-tests do it first)
.\scripts\run.ps1 extract         # Python path without VBA: Jira -> jira2dashboard.xlsx (values + SVG charts)
.\scripts\manual-test.ps1         # manual try: fake Jira + pre-filled copy (mock-workbook), opened in a new Excel
.\scripts\stop-mock.ps1           # kill the fake Jira that start-mock.ps1 / manual-test.ps1 started
.\scripts\release.ps1 -Version v1.0.0 [-Publish]   # zip of HEAD + built workbook in out\release; -Publish = GitHub release
```

The release zip leaves out the `export-ignore` paths of `.gitattributes` (CLAUDE.md, the handoff, `archive/`,
`.claude/`, `.venv/`); `tests/unit/test_release_contents.py` pins the list. `-Publish` is outward-facing: ask AL first.

`test` and `coverage` skip the `excel` and `network` markers by default (`addopts` in `pyproject.toml`): use `test-excel` /
`test-network` for those.

Every script and command answers `--help` (`tests/unit/test_help.py` checks it). A new Python entry point parses
its arguments first with `cli/args.parser` (the module docstring is the help), so `--help` runs nothing. A new
`.ps1` gets comment-based help (`.SYNOPSIS`...) and `[CmdletBinding(PositionalBinding = $false)]` with a `$Help`
switch and a `ValueFromRemainingArguments` `$Rest` (so `--help` lands in `$Rest`, never in `-Port`); `-?` works in
a PowerShell console, not through `powershell -File`.

Single test: `.\scripts\run.ps1 test tests/unit/test_checks.py::test_name` or
`.\scripts\run.ps1 test-excel tests/steps/test_flow_excel.py -k "contr" -q`. Call `run.ps1` directly from PowerShell: do not
wrap it in `cmd /c` or redirect with `2>&1` (it stops on `$ErrorActionPreference='Stop'`).
From Bash: `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/run.ps1 test-excel > "$TEMP/x.log" 2>&1` works
(the redirect is outside PowerShell). Write logs, benches and scratch files to `$TEMP` or the scratchpad, never to
`out/`. `out\release` holds what is delivered (the workbook of `build`, the packages of `release.ps1`) and is never
cleaned; `out\manual_test` holds the manual-test copy `jira2dashboard4manualtest.xlsm` (and its `-Python` xlsx),
never cleaned either; `out\inter` holds the rest (test workbooks, exports, `.coverage`, `.pytest_cache`),
emptied by `build` / `vba-tests` except the caches (`src/jira2dashboard/clean.py`). Tests write to `INTER_DIR`.

Before finishing a change run the full pass: `coverage`, `vba-tests`, `test-excel`, `test-network`, `build`.

## Architecture (big picture)

- **Python generates, Excel computes.** `src/jira2dashboard/build_workbook.py` writes the parameter sheets, the empty raw
  sheets and a *plan*, injects the `vba/*.bas` modules and compiles them. VBA writes only raw data (`Issues`,
  `StatusIntervals`); **every figure is an Excel formula** in a table (`ItemFlow`, `DailyLoad`, `Metrics`, `DataChecks`).
- **Formulas and charts are data.** `flow/formulas.py` and `flow/dashboard_spec.py` hold the formula text and chart
  descriptions (pure, tested without Excel). `flow/blueprint.py` writes them into the sheets `FlowFormulas`
  (Table|Column|Formula; empty Formula = a values column), `ChartSpec` and `ChecksSpec`. After each extraction VBA
  `modFlowBuild.BuildFlowSheetsIn` rebuilds all tables, formulas and charts from those sheets, sized to the new data.
  Add a metric or check by editing the plan, not VBA.
- **Python reference as oracle.** `flow/` (`intervals.py`, `itemflow.py`, `load.py`, `metrics.py`, `checks.py`) is a pure
  Python implementation of the same metrics and data checks. The scenarios in `tests/features/flow_excel.feature` run the
  real delivered workbook against the fake Jira and compare Excel with this reference **cell by cell** (Flow Load modes A
  and B). Keep the two sides in step: a formula change needs the reference change and vice versa.
- **Parameters live in the workbook, never in code.** `params/*.md` (config, projects, workflows, transitions,
  field_mapping, periods) are read and validated by `src/jira2dashboard/params/`, written into sheets, and read by VBA
  (`modParams`). Status names, team names and extract columns must not appear as literals in VBA (a guard test enforces it).
  `Config` columns: Parameter | Value | Example | Mandatory | Description (formulas read A and B; the rest by header). An empty Mandatory=True value
  stops the buttons before any call to Jira (`modParams.RequireMandatory`, error 608).
- **User manual.** `src/jira2dashboard/manual.py` (pure text) is written by `flow/sheets.write_manual` at the delivery
  build: the first steps on the Dashboard from T1, and the sheet `Manual`. Update it when a button, sheet or step changes.
- **Extraction (VBA).** `modHttp` (WinHTTP, proxy with Windows auth, retries; `modProxy` finds the proxy of the Windows
  settings when `ProxyServer` is empty: WPAD, PAC or fixed, via WinHTTP API calls on byte buffers, not VBA Types, which
  crashed 64-bit Excel) -> `modJira` (search, bulk changelog) ->
  `modExtract` -> `modFlow` (status intervals) ; `modActions` is the dashboard button (`ExtractAndCompute`). The previous
  results are cleared only after Jira has answered. `modLog` writes the local log (`%LOCALAPPDATA%\jira2dashboard\logs`):
  RUN, BEGIN/END per step, ERROR; never pass it a token, URL, JQL or Jira answer (`test_vba_log_guard.py` checks).
- **Python path without VBA (`standalone/`, `run.ps1 extract`).** A second consumer of the flow reference: it reads
  the parameter sheets of the filled workbook as a zip (`xlsx_read`, by header name, shape rules shared with
  params/*.md through `params.model.shape_problems`, Config rules of the VBA in `params_book`), calls Jira with the
  WinHTTP COM object (`http_winhttp`: same proxy, Windows auth and retry rules as `modHttp`; body read from
  ResponseBody as UTF-8; injectable sleep) after `proxy.py` (ctypes port of `modProxy`), maps the JSON like
  `modExtract`/`modFlow` (`mapping`, oracle: `test_mapping` against `flow_adapter`), computes with `flow/`
  (`results`, same column names as `formulas.TABLES`), draws the 8 charts as SVG (`charts_svg`) and writes values +
  pictures with zipfile (`xlsx_write`: SVG with a blank PNG fallback, Excel 365 shows msoGraphic). Keep it in step
  with the VBA when an extraction rule changes. Its log (`python.log`, same folder as the VBA log) is complete on
  purpose (AL's choice): status, call, Jira's message, tracebacks; never the token nor the Authorization header
  (`Scrubbed` formatter; `test_extract_main` checks).
- **Jira configuration (VBA).** `modJiraConfig` (button Read the Jira configuration, `BtnReadJiraConfiguration`) reads
  projects, issue types, statuses, workflows (`workflows/preview`) and fields into proposal sheets `Jira*`; it never writes
  the parameter sheets (calls and mapping: handoff section 8c). Its Rovo counterpart, to compare Jira with the workbook:
  `rovo/rovo_settings_analyser.md` (agent config, not code).
- **Rovo side (md only, no code).** `rovo/` holds the master copies of the Rovo Studio agents. `rovo/rovo-guide.md` is
  the current guide; `docs/rovo-guide.md` is frozen. Steering is hybrid: the workbook computes, and the RTE publishes
  the figures on the page `rovo/confluence_flow_metrics_page.md`. The agent `rovo/rovo_train_steering.md` reads that
  page and never computes a metric; its skills are in `rovo/skills/`. The page template copies the column names of
  `METRICS_COLUMNS` and `DATACHECKS_COLUMNS` and the chart titles: update it when they change.
  Label every Rovo fact Documented / Community / Verify in UI / Proposal; the check coverage of `rovo/rovo-guide.md`
  section 3 (10 exact, 9 partial, 1 to test, 9 workbook only) is the reference for the other `rovo/` files. Research
  behind them: `C:\dev\reports\Agents Rovo Studio et métriques.md` (outside the repo).
- **Fake Jira.** `mockjira/` is a stdlib `http.server` Jira Cloud (search/jql, changelog bulkfetch, proxy/CONNECT, fault
  injection) fed by `dataset.py` (hand-written cases via `issue()`) and `generator.py` (seeded volume);
  `anomalies.py` adds known data-quality cases; `config.py` answers the configuration calls from the parameters (round
  trip). The real Jira is never called. Manual try: `scripts\manual-test.ps1` = `start-mock.ps1` (mock with
  `--anomalies`, pid + start time in `%TEMP%\jira2dashboard`) + `mock_workbook.py` (the copy
  `out\manual_test\jira2dashboard4manualtest.xlsm`, Config, Periods, PIs filled; the source stays empty) opened in
  its own Excel (`/x`); `-Python` runs `extract` on the copy instead.
  `stop-mock.ps1` kills only that pid with that start time (`taskkill /T`), never a port owner it did not start.
  Limit: with `proxy_challenge`, the fake proxy accepts the first NTLM leg, which WinHTTP sends with an empty body,
  so a POST through it arrives empty; the Windows session sent to a proxy is tested with GET only.
  Its request log stores header names in lower case (`proxy-authorization`).
- **COM layer.** `excel/com.py` holds the pure helpers and `excel_session`; `testbench/` builds, compiles and runs the VBA
  tests; `testbench/children/` scripts run as separate processes by path. Files omitted from coverage because they need
  COM: `flow/sheets.py`, `cli/check_machine.py`, `mockjira/__main__.py`, and the `build*`/`main` functions (see
  `pyproject.toml`).

Domain decisions and the metric definitions are in `docs/` (7 md files, **never edit them**: their MD5 is tracked in
`claude-code-handoff.md`, which also holds the roadmap, review dispositions and the MD5 of every `.bas`). Update the MD5
table in the handoff whenever a `.bas` changes. `README.md` (English) is the engineering guide: architecture,
scripts, configuration, client use, logs and debugging.

## Adding a data check

1. `flow/checks.py`: a `Check` row and the rule in `issue_checks` (Python reference).
2. `flow/formulas.py`: a `DCnn` column in `ITEMFLOW_COLUMNS` (helper columns in `INTERVAL_HELPERS` if needed).
3. If it needs a new Jira field: a raw column in `ISSUES`, `ExtractFields` and the row in `modExtract.bas`, the fixtures in
   `modTests.bas`, and the field in `mockjira/` (`FlowIssue` and `flow_adapter.py`).
4. `mockjira/anomalies.py`: a known case with the ids it must trigger; the data-check scenario of `flow_excel.feature`
   compares it with the workbook.
5. Update the counts (marker test, the Config row count in `params.feature` if a parameter is added), the MD5 table and
   the handoff.

## Development practices

- **BDD outer loop, TDD inner loop.** Start from a Gherkin scenario (`tests/features/*.feature`, French keywords,
  `# language: fr`, steps in `tests/steps/`, pytest-bdd), then drive the code with unit tests (`tests/unit/`, pure Python,
  fast) in red-green-refactor steps. A feature is done when its scenario passes on the real workbook, not only its unit tests.
- **Craftsmanship.** KISS, YAGNI, DDD vocabulary (`flow/model.py`: Issue, Interval, ItemFlow, Rules...). Keep logic in pure
  modules and push COM/IO to thin edges so it can be tested without Excel. Prove that a new test discriminates: break the
  code on purpose (mutation) and watch it fail, then restore.
- **Coverage.** Pure Python >= 90 % (currently about 99 %), enforced by `run.ps1 coverage`. VBA branches carry
  `Cov "ID"` markers; `vba-tests` requires **100 % of markers hit** and the unit test
  `test_real_vba_modules_are_ascii_and_have_unique_markers` pins the marker count and ranges: update it when adding markers.
  Trivial entry points (`modActions`, `modButtons`) carry none.
- **VBA conventions.** ASCII only, LF line endings, `Option Explicit`; offline tests `T_*` and fake-Jira tests `M_*` in
  `modTests.bas`; button macros are Public Subs without arguments.
- **VBA traps.** A one-cell `Range.Value` is not an array; assigning an object to a Variant needs `Set` (error 450);
  `Application.OnTime` has a 1 s granularity; a COM call to an Excel that runs a macro is rejected (see `retry()` in
  `tests/steps/test_flow_excel.py`).
- **Test markers.** `excel` and `network` tests are excluded from the default run; Excel scenarios build the delivered
  workbook once per module and open it in a fresh Excel per scenario.
- **pytest-bdd binding.** Steps are local to their module: a new `.feature` that reuses Excel steps is bound with a second
  `scenarios()` call in `tests/steps/test_flow_excel.py`. A `.feature` bound by no `scenarios()` call is not run (used for
  scenarios written ahead, e.g. `workflow_by_type.feature`).
- **`M_` tests** (VBA against the fake Jira) run only in the `@excel` scenario of `mock_jira.feature`, not in `vba-tests`
  (`T_` and `N_` only). `MockReset` clears the fake Jira's request log, which that scenario reads after the proxy tests:
  keep tests that reset before `M_ThroughProxy` in `RunMockTests`.

## Hard constraints

- **Nothing from the client may leave the client machine** (no scripts, no data). Client values (`BaseUrl`, `Email`,
  `ApiToken`, `ProxyServer`, `ProxyBypass`, `Filter`, `WorkTypeField`, `AsOf`, the periods calendar) stay **empty** in the
  repo; the API token must never be written to `params/config.md` (a test refuses it).
- **Never kill an `EXCEL.EXE` the scripts did not start** (the user may have their own), except one hung for long on a
  VBA error dialog (AL allows it). A modal VB dialog (an untrapped VBA error) is the usual cause of a hung Excel run.
- **Do not call a real Jira.** Use the fake Jira.
- Excel formula traps: a formula typed in one cell of a table column fills the whole column (write key columns as values
  before creating the table); a structured reference to a column whose name contains `_` is refused; compare dates with a
  half-second margin; a blank result is `""`, never `0`.
- Mind the COM traps listed in `claude-code-handoff.md` section 7: positional arguments only, `put_en_us` for number formats,
  leading apostrophe on text written to cells (`as_cell` / `SafeCell`), read `.Value2`, activate a chart's sheet before export.
- **Git.** Repo `AlexisAIchangeAgent/jira2dashboard` (public), branch `main`. `.gitattributes` disables line-ending
  conversion (`* -text`, `*.bas text eol=lf`) so the tracked MD5 hold; Excel binaries (`*.xlsm`, `out/`) are ignored.
  Commit messages end with the `Co-Authored-By` line of the model in use (e.g. `Claude Opus 5.5 <noreply@anthropic.com>`). If SSH push is refused, push over
  HTTPS with the `gh` token: `git -c credential.helper= -c "credential.helper=!gh auth git-credential" push https://github.com/AlexisAIchangeAgent/jira2dashboard.git main`.

## Environment gotchas

- **Excel vanishes in the Excel scenarios ("RPC server unavailable")**: watch the `EXCEL.EXE` exit code. `0xC0000005` is
  a crash; Office's crash recovery can also end it with 0 and start a new `EXCEL.EXE` in the same millisecond (do not
  read that as an Office update). The Application event log (Application Error, event 1000) names the faulting module:
  ask AL before reading it. It also hits `vba-tests`. Before blaming a change for a flaky Excel run, run the last commit
  the same number of times: `git worktree add --detach <scratchpad>\wt HEAD`, then pytest there with the main `.venv`
  python and `PYTHONPATH=<wt>\src` (`git stash` leaves new untracked files in place). With a failure rate near 40 %, a
  fix is proven only by 5 green module runs in a row.
- **Open defect:** Excel crashes (`0xC0000005`, a DEP fault inside VBA's engine, module "unknown") during the
  extraction, always right after the last search page (`END search page 30: 150 issues` in the local log). The button
  scenario alone crashed 7 times out of 10 on 2026-11-05 (`test-excel tests/steps/test_flow_excel.py -k bouton`): that is
  the quickest repro. It predates the 2026-11-04 review. Read handoff section 8b before chasing it.
- **Git Bash rewrites `/switch` arguments** of Windows programs into paths (`/update` -> `C:/.../Git/update`, also
  `wevtutil /q:`): prefix the command with `MSYS_NO_PATHCONV=1`.
- The PowerShell tool is slow here (CIM / event-log queries time out): use Bash `tasklist` for processes.
- Reading PowerShell output from Python (`Get-Help`, the scripts): `encoding="oem", errors="replace"`, not UTF-8.
- Edit scripts with quotes, `\r\n` or backslashes (Windows paths in Markdown): write a `.py` file with the Write tool and
  run it, not a Bash heredoc (a heredoc turned `tests\\features` into a form feed in the handoff).
  Pattern: `assert s.count(old) == 1` before each replace, write with `newline="\n"`, run with `.venv/Scripts/python.exe`.
- `Windows fatal exception: code 0x800706be` (or `0x800706ba`) printed after Excel scenarios is teardown noise, also in
  green runs: only the test result counts.

## Roadmap not done yet (ask before starting)

Real test at the client machine is still to be done by AL.
Pending: a workflow per issue type (scenarios written, not bound; bind them once the client confirms the need); the Rovo
agents to try at the client: Settings Analyser (its section 5) and ART Flow Steering (its section 6). For the second,
also check whether a custom skill can be attached to a Studio agent.
