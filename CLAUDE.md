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
.\scripts\run.ps1 coverage        # same + coverage of the pure code; fails under 93 %
.\scripts\run.ps1 test-excel      # scenarios that drive Excel (marker `excel`, no network)
.\scripts\run.ps1 test-network    # Excel + network scenarios (marker `network`)
.\scripts\run.ps1 vba-tests       # build + compile + run the VBA test suite; --offline, --bisect, --out name.xlsm
.\scripts\run.ps1 check-machine   # 22 checks of the machine (Excel, VBA trust, macro of the workbook)
.\scripts\run.ps1 diagnose        # paste-ready report when the macros are blocked (hashes, Mark of the Web, Office, security products)
.\scripts\run.ps1 mock-jira       # fake Jira on http://127.0.0.1:8099 (--anomalies: with the known data anomalies)
.\scripts\run.ps1 build           # delivery workbook out\release\jira2dashboard.xlsm (--out name.xlsm: to out\inter)
.\scripts\run.ps1 clean           # empty out\inter except the test caches (build and vba-tests do it first)
.\scripts\run.ps1 extract         # Python path without VBA: Jira -> jira2dashboard.xlsx (values + SVG charts)
.\scripts\manual-test.ps1         # manual try: fake Jira + pre-filled copy (mock-workbook), opened in a new Excel
.\scripts\stop-mock.ps1           # kill the fake Jira that start-mock.ps1 / manual-test.ps1 started
.\scripts\scenario.ps1            # unattended scenario: fake Jira, filled copy, Python extraction, checks, clean-up (-Excel: + test-excel)
.\scripts\run.ps1 test-scripts    # Pester 3.4 tests of scripts\*.ps1 (tests\ps), in process, with fakes; fails under 93 % command coverage
.\scripts\release.ps1 -Version v1.0.0 [-Publish]   # zip of HEAD + built workbook in out\release; -Publish = GitHub release
```

Every `scripts\*.ps1` has a double-click launcher `scripts\<name>.cmd` (shared helper `_launch.cmd`: `-ExecutionPolicy Bypass`
for that run only, waits for a key only on a double-click or with `J2D_PAUSE=1`, `--help`, exit code kept). A new `.ps1` needs
one: copy `stop-mock.cmd` and change its `set` lines. `.cmd` files are ASCII with CRLF and their description line has no
`<>|&^%!` (cmd would take them for operators); `tests/unit/test_cmd_launchers.py` pins all this. From Git Bash call them with
`MSYS_NO_PATHCONV=1 cmd.exe /c .\scripts\x.cmd`: otherwise `/c` becomes a path and an interactive cmd opens.

The release zip leaves out the `export-ignore` paths of `.gitattributes` (CLAUDE.md, `archive/`,
`.claude/`, `.venv/`); `tests/unit/test_release_contents.py` pins the list. `-Publish` is outward-facing: ask AL first.
Releases: the version is `vX.Y.Z` only (`v0.0.0-dryrun` is refused) and the number is AL's choice. The dry run
(`release.ps1 -Version v0.0.0`) needs a clean tree, builds with Excel and leaves `out\release\<version>`: delete it afterwards.
`-Publish` also needs HEAD == origin/main and `gh` logged in. `v0.1.0` (2026-11-06), `v0.2.0` (2026-11-07, commit 878807e) and `v0.3.0` (2026-11-11, commit 31c9574) are published.
`diagnose` never opens the workbook, never reads `python.log` (JQL, Jira's messages) nor the Windows event log, and masks the home
folder: keep it that way when adding a section. Allow-lists of an endpoint protection ask for the MD5 (`notes.md` has it) or, lasting longer, the path and name:
the workbook is always `jira2dashboard.xlsm` (`paths.DELIVERED_NAME`), never with the version in its name.

`test` and `coverage` skip the `excel` and `network` markers by default (`addopts` in `pyproject.toml`): use `test-excel` /
`test-network` for those.

Every script and command answers `--help` (and every `.ps1` also `-Help`; `tests/unit/test_help.py` checks each
file of `scripts\`). A new Python entry point parses its arguments first with `cli/args.parser` (the module docstring is the help), so `--help` runs nothing. A new
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

Before finishing a change run the full pass: `coverage`, `vba-tests`, `test-excel`, `test-network`, `build`
(skill `/full-pass`). The Claude config is versioned in `.claude/` (like this file) but left out of the release
(`export-ignore`); only `.claude/settings.local.json`, the personal settings, is gitignored. A PostToolUse hook runs the
VBA guard tests after each edit of `vba/*.bas` (`.claude/hooks/bas_guard.py`); skills `full-pass`, `add-data-check`.

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
  field_mapping, iterations) are read and validated by `src/jira2dashboard/params/`, written into sheets, and read by VBA
  (`modParams`). Status names, team names and extract columns must not appear as literals in VBA (a guard test enforces it).
  `Config` columns: Parameter | Value | Example | Mandatory | Description (formulas read A and B; the rest by header). An empty Mandatory=True value
  stops the buttons before any call to Jira (`modParams.RequireMandatory`, error 608).
- **User manual.** `src/jira2dashboard/manual.py` (pure text) is written by `flow/sheets.write_manual` at the delivery
  build: the first steps on the Dashboard from T1, and the sheet `Manual`. Update it when a button, sheet or step changes.
- **Extraction (VBA).** `modHttp` (WinHTTP, proxy with Windows auth, retries; `modProxy` finds the proxy of the Windows
  settings when `ProxyServer` is empty: WPAD, PAC or fixed, via WinHTTP API calls on byte buffers, not VBA Types, which
  crashed 64-bit Excel) -> `modJira` (search, bulk changelog) ->
  `modExtract` -> `modFlow` (status intervals) ; `modActions` is the dashboard button (`ExtractAndCompute`). The previous
  results are cleared only after Jira has answered. `modLog` writes the local log in the folder `logs` next to the workbook (it runs from a folder not known in advance; `%LOCALAPPDATA%\jira2dashboard\logs` is the fallback):
  RUN, BEGIN/END per step, ERROR; never pass it a token, URL, JQL or Jira answer (`test_vba_log_guard.py` checks).
- **Python path without VBA (`standalone/`, `run.ps1 extract`).** A second consumer of the flow reference: it reads
  the parameter sheets of the filled workbook as a zip (`xlsx_read`, by header name, shape rules shared with
  params/*.md through `params.model.shape_problems`, Config rules of the VBA in `params_book`), calls Jira with the
  WinHTTP COM object (`http_winhttp`: same proxy, Windows auth and retry rules as `modHttp`; body read from
  ResponseBody as UTF-8; injectable sleep) after `proxy.py` (ctypes port of `modProxy`), maps the JSON like
  `modExtract`/`modFlow` (`mapping`, oracle: `test_mapping` against `flow_adapter`), computes with `flow/`
  (`results`, same column names as `formulas.TABLES`), draws the 8 charts as SVG (`charts_svg`) and writes values +
  pictures with zipfile (`xlsx_write`: SVG with a blank PNG fallback, Excel 365 shows msoGraphic). Keep it in step
  with the VBA when an extraction rule changes. Its log (`logs\python.log` next to the workbook, as the VBA log) is complete on
  purpose (AL's choice): status, call, Jira's message, tracebacks; never the token nor the Authorization header
  (`Scrubbed` formatter; `test_extract_main` checks).
- **Support codes.** Every error shown or logged carries `J2D-<AREA>-<NNN>` (`CFG`, `JIRA`, `NET`, `DATA`, `FLOW`, `ROVO`, `SYS`, `VBA`, `GEN`):
  the table is `src/jira2dashboard/error_codes.py` (code, vba number, meaning, cause, check, what support asks), shown in the sheet
  `ErrorCodes`; `modLog.ErrorCode(number, description)` is the VBA rule (602 has two causes: BaseUrl check, HTTP status = `JIRA-<status>`,
  status 0 = `NET-<WinHTTP error>`), `error_codes.code_of_vba` and `standalone.support_code` the Python one. A new `Err.Raise vbObjectError + n`
  needs its `Case n` in `ErrorCode` and a row in the table (`test_error_codes.py` fails otherwise); a code is never reused or renumbered.
  The scripts use `J2D-SCR-nnn` (every `Fail`, `throw` and unknown-argument message of `scripts\*.ps1` carries one: a test pins it), `diagnose` lists
  its `Findings` as `J2D-SYS-02x` with the line `Codes to quote to the support`, `check-machine` prints `J2D-SYS-040` on a FAIL.
  `test-excel` keeps the VBA log of each scenario in `scripts\logs\excel-tests\` (`testbench/keep_log.py`): read it after a run.
- **Jira configuration (VBA).** `modJiraConfig` (button Read the Jira configuration, `BtnReadJiraConfiguration`) reads
  projects, issue types, statuses, workflows (`workflows/preview`) and fields into proposal sheets `Jira*`; it never writes
  the parameter sheets (calls and mapping: the header of `modJiraConfig.bas`). Besides the five first sheets it writes
  `JiraProjectWorkflows`, `JiraProjectFields` (createmeta, needs Create issues) and `JiraFieldUsage` (one page of issues per project). Its Rovo counterpart, to compare Jira with the workbook:
  `rovo/rovo_settings_analyser.md` (agent config, not code).
- **Prompt for Rovo (button, VBA).** `src/jira2dashboard/rovo_prompt.py` (pure text) is written at build into the sheet
  `RovoPromptSpec` (Part | Name | Text); `modParams.WriteRovoPrompt` (button `modActions.BtnRovoPrompt`, cell `R1`) turns it into the
  sheet `RovoPrompt`: Config read live (name, mandatory, meaning, example, never the Value), each sheet with its header, and the CSV
  format asked of Rovo (Sheet, Row, Column, Value, Note). No clipboard call (the VBA creates only two COM objects): the lines are selected.
  The button Load Rovo CSV (`BtnLoadRovoCsv`, cell `R3`; `ImportRovoCsvFile(path)` is its work without the dialogs) reads Rovo's CSV
  (`modParams.ParseCsv`, `RovoCsvRead`, `RovoCsvApply`: comma, semicolon or tab, UTF-8 with or without BOM): Config gets the Value
  of each parameter named, every other sheet of the file has all its rows replaced; UNKNOWN, empty and the token placeholder are
  not loaded.
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
  `out\manual_test\jira2dashboard4manualtest.xlsm`, Config, Iterations, PIs filled; the source stays empty) opened in
  its own Excel (`/x`); `-Python` runs `extract` on the copy instead.
  `stop-mock.ps1` kills only that pid with that start time (`taskkill /T`), never a port owner it did not start.
  Limit: with `proxy_challenge`, the fake proxy accepts the first NTLM leg, which WinHTTP sends with an empty body,
  so a POST through it arrives empty; the Windows session sent to a proxy is tested with GET only.
  Its request log stores header names in lower case (`proxy-authorization`).
- **COM layer.** `excel/com.py` holds the pure helpers and `excel_session`; `testbench/` builds, compiles and runs the VBA
  tests; `testbench/children/` scripts run as separate processes by path. Files omitted from coverage because they need
  COM: `flow/sheets.py`, `cli/check_machine.py`, `mockjira/__main__.py`, and the `build*`/`main` functions (see
  `pyproject.toml`).

Domain decisions and the metric definitions are in `docs/` (7 md files, **never edit them**). The handoff file `claude-code-handoff.md` was deleted
on 2026-11-06; its last version is `git show 6a93cdd:claude-code-handoff.md` (section 7: COM traps, 8b: open defect, 8c: Jira calls, and the roadmap).
`README.md` (English) is the engineering guide: keep it concise,
tables rather than code blocks for lists, never the word "client" nor the vendor or product name of the endpoint protection
(public repo: say EPDR or "endpoint protection"), prose lines <= 120 characters, ASCII for the architecture,
Mermaid for the workflows. Check a Mermaid edit with the real parser (`npm i mermaid jsdom` in the scratchpad,
then `mermaid.parse`), and check every `](#anchor)` link against the headings.

## Adding a data check

1. `flow/checks.py`: a `Check` row in `CHECKS` and the rule in `run_checks` (`_status_history`, `_field_checks`,
   `_link_flags` or `fill_rates`): the Python reference.
2. `flow/check_formulas.py`: the `DCnn` entry of `FLAGS`, which `ITEMFLOW_COLUMNS` turns into an `ItemFlow` column
   (helper columns in `HELPERS`, or `INTERVAL_HELPERS` of `flow/formulas.py`). Skill: `/add-data-check`.
3. If it needs a new Jira field: a raw column in `ISSUES`, `ExtractFields` and the row in `modExtract.bas`, the fixtures in
   `modTests.bas`, and the field in `mockjira/` (`FlowIssue` and `flow_adapter.py`).
4. `mockjira/anomalies.py`: a known case with the ids it must trigger; the data-check scenario of `flow_excel.feature`
   compares it with the workbook.
5. Update the counts (marker test, the Config row count in `params.feature` if a parameter is added).

## Development practices

- **BDD outer loop, TDD inner loop.** Start from a Gherkin scenario (`tests/features/*.feature`, French keywords,
  `# language: fr`, steps in `tests/steps/`, pytest-bdd), then drive the code with unit tests (`tests/unit/`, pure Python,
  fast) in red-green-refactor steps. A feature is done when its scenario passes on the real workbook, not only its unit tests.
- **Craftsmanship.** KISS, YAGNI, DDD vocabulary (`flow/model.py`: Issue, Interval, ItemFlow, Rules...). Keep logic in pure
  modules and push COM/IO to thin edges so it can be tested without Excel. Prove that a new test discriminates: break the
  code on purpose (mutation) and watch it fail, then restore.
- **Coverage.** Pure Python >= 93 % (currently about 99 %), enforced by `run.ps1 coverage`. VBA branches carry
  `Cov "ID"` markers; `vba-tests` requires **100 % of markers hit** and the unit test
  `test_real_vba_modules_are_ascii_and_have_unique_markers` pins the marker count and ranges: update it when adding markers.
  Trivial entry points (`modActions`, `modButtons`) carry none.
- **VBA conventions.** ASCII only, LF line endings, `Option Explicit`; offline tests `T_*` and fake-Jira tests `M_*` in
  `modTests.bas`; button macros are Public Subs without arguments.
- **VBA traps.** A one-cell `Range.Value` is not an array; assigning an object to a Variant needs `Set` (error 450);
  `Application.OnTime` has a 1 s granularity; a COM call to an Excel that runs a macro is rejected (see `retry()` in
  `tests/steps/test_flow_excel.py`).
- **VBA speed traps** (measured on the dev machine, 2026-11-07: the full extraction of the fake Jira, 1 244 issues, went from
  245 s to 24 s). `TypeName(object)` costs about 0.3 ms: use `TypeOf x Is Collection` or an `On Error` probe, never in a loop
  (`JGet` was 300 us, now 2 us). The first call on a fresh `Scripting.Dictionary` costs about 0.65 ms, even `Exists`
  (a `Collection` costs nothing), and `d.CompareMode = 0` on a new one about 0.3 ms: the default is already 0. Each JSON
  object is a Dictionary, so `JsonParse` still takes about 15 s of the 24 (a lighter JSON object is the next lever). `COUNTIFS`
  on a whole column (`Sheet!$A:$A`) scans its 1 048 576 rows (10 s for 7 284 cells; `SUMIFS`/`MATCH` stop at the used
  range, `COUNTIFS` does not): cut it at the last row with `refs.bounded` and a control cell (`test_no_countifs_reads_a_whole_column`).
  A step of the extraction job must stay short (about 1 s): `CHANGELOG_CHUNK` (50 issues) and `ROW_BATCH` (100 issues; a test
  sets `gRowBatch`). Profile with the log (`BEGIN`/`END` of each sub-step) and a scratch copy of the workbook that gets a
  temporary benchmark module through `VBProject.VBComponents.Add` (never commit it).
- **Test markers.** `excel` and `network` tests are excluded from the default run; Excel scenarios build the delivered
  workbook once per module and open it in a fresh Excel per scenario.
- **pytest-bdd binding.** Steps are local to their module: a new `.feature` that reuses Excel steps is bound with a second
  `scenarios()` call in `tests/steps/test_flow_excel.py`. A `.feature` bound by no `scenarios()` call is not run (used for
  scenarios written ahead, e.g. `workflow_by_type.feature`).
- **Test hooks in VBA:** a wait or a failure that tests cannot trigger gets a `Public` flag read behind a `Cov` marker
  (`gHttpFake`, `gHttpCancelInWait` in `modHttp`). `modActions` has no markers: only `flow_excel.feature` covers it.
- **`M_` tests** (VBA against the fake Jira) run only in the `@excel` scenario of `mock_jira.feature`, not in `vba-tests`
  (`T_` and `N_` only). `MockReset` clears the fake Jira's request log, which that scenario reads after the proxy tests:
  keep tests that reset before `M_ThroughProxy` in `RunMockTests`.

## Hard constraints

- **Nothing from the client may leave the client machine** (no scripts, no data). Client values (`BaseUrl`, `Email`,
  `ApiToken`, `ProxyServer`, `ProxyBypass`, `Filter`, `WorkTypeField`, `AsOf`, the iterations calendar) stay **empty** in the
  repo; the API token must never be written to `params/config.md` (a test refuses it).
- **Never kill an `EXCEL.EXE` the scripts did not start** (the user may have their own), except one hung for long on a
  VBA error dialog (AL allows it). A modal VB dialog (an untrapped VBA error) is the usual cause of a hung Excel run.
- **Do not call a real Jira.** Use the fake Jira.
- **The delivered VBA creates only `WinHttp.WinHttpRequest.5.1` and `Scripting.Dictionary`** (`tests/unit/test_delivered_vba_com_objects.py`): an endpoint
  protection reads `ADODB.Stream` (next to an HTTP client) and `MSXML2` (as a Base64 codec) as a downloader. UTF-8 and Base64 are plain VBA
  in `modText`; `modTests` keeps ADODB and MSXML only as the oracle of `T_Utf8Oracle` and `T_Base64Oracle`.
- Excel formula traps: a formula typed in one cell of a table column fills the whole column (write key columns as values
  before creating the table); a structured reference to a column whose name contains `_` is refused; compare dates with a
  half-second margin; a blank result is `""`, never `0`.
- Mind the COM traps (section 7 of the deleted handoff, see above): positional arguments only, `put_en_us` for number formats,
  leading apostrophe on text written to cells (`as_cell` / `SafeCell`), read `.Value2`, activate a chart's sheet before export.
- **Git.** Repo `AlexisAIchangeAgent/jira2dashboard` (public), branch `main`. `.gitattributes` disables line-ending
  conversion (`* -text`, `*.bas text eol=lf`) so the tracked MD5 hold; Excel binaries (`*.xlsm`, `out/`) are ignored.
  Commit messages end with the `Co-Authored-By` line of the model in use (e.g. `Claude Opus 5.5 <noreply@anthropic.com>`). If SSH push is refused, push over
  HTTPS with the `gh` token: `git -c credential.helper= -c "credential.helper=!gh auth git-credential" push https://github.com/AlexisAIchangeAgent/jira2dashboard.git main`.
  A push refused with `fatal error in commit_refs_rules` is a GitHub-side glitch (the repo has no ruleset): retry
  a few minutes later, do not chase it.

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
  the quickest repro. It predates the 2026-11-04 review. Read section 8b of the deleted handoff (see above) before chasing it.
- **Git Bash rewrites `/switch` arguments** of Windows programs into paths (`/update` -> `C:/.../Git/update`, also
  `wevtutil /q:`): prefix the command with `MSYS_NO_PATHCONV=1`.
- The PowerShell tool is slow here (CIM / event-log queries time out): use Bash `tasklist` for processes.
- Reading PowerShell output from Python (`Get-Help`, the scripts): `encoding="oem", errors="replace"`, not UTF-8.
- Edit scripts with quotes, `\r\n` or backslashes (Windows paths in Markdown): write a `.py` file with the Write tool and
  run it, not a Bash heredoc (a heredoc turned `tests\\features` into a form feed in a Markdown file).
  Pattern: `assert s.count(old) == 1` before each replace, write with `newline="\n"`, run with `.venv/Scripts/python.exe`.
- `Windows fatal exception: code 0x800706be` (or `0x800706ba`) printed after Excel scenarios is teardown noise, also in
  green runs: only the test result counts.
- **Hung test Excel (VBA error dialog):** run `vba-tests` with `timeout 420`. The Excel to kill is the one whose window
  title is "Microsoft Visual Basic" and whose start time matches the run (`Get-Process EXCEL | Select Id, StartTime,
  MainWindowTitle`), never an older one. Mutation proof of a VBA test: copy the module to `$TEMP`, mutate it with a Python
  script (`assert count == 1`), run, restore with `cp` in the same command. A new check that reads Excel objects must not
  raise (a raise opens that dialog): wrap it like `AxisEnds` in `modTests.bas`.
- **Excel hygiene (AL's rule):** after every Excel run, even a green one, `tasklist | grep -c EXCEL` must be back to what it was;
  stop only the Excel your run started (start time; title "Microsoft Visual Basic" when hung), never an older one.
- `com_error -2146959355` ("Echec de l'execution du serveur") at `DispatchEx` = Excel failed to start: transient, rerun.
- **A few VBA tests only:** copy `modTests.bas`, cut `TestNames()` to the tests wanted, run `vba-tests --offline`, restore with `cp`
  (about 1 minute instead of 5; never commit the cut list).
- `_launch.cmd` traps: a `for` list drops `/?` (wildcard), so it is tested apart; a launcher starts the helper without `call`
  (`call` swallows `/?`); from a test, `cmd /c` needs `.\name.cmd`.
- After moving code between `.bas` modules: `Private` names (Const, Function) used from the new module break the compile of
  the test workbook only (`build` compiles the delivered modules, `vba-tests` adds `modTests`). `--bisect` loads the
  modules in order, so a forward reference shows as a false failure at an early module: scan for `Private` names instead.
  Python steps call `modX.Name` explicitly (`tests/steps/test_params.py`): update them.
- The local `.claude` hook runs only after Edit/Write: a `.bas` changed through a Bash/Python script skips it. Run the 3
  guard tests by hand.
- **Checked facts (Microsoft Learn MCP, `mcp__claude_ai_Microsoft_Learn__*`):** entering a formula in manual calculation
  calculates it, so `CalculateFull` after the switch back to automatic is needed only when Excel was manual;
  `WinHttpSetTimeouts` does not bound proxy auto-detection (WPAD); `Application.StatusBar` reads `False` or the text
  "FALSE" (compare with `UCase$(CStr(...))`).

## Roadmap not done yet (ask before starting)

Real test at the client machine is still to be done by AL.
Pending: a workflow per issue type (scenarios written, not bound; bind them once the client confirms the need); the Rovo
agents to try at the client: Settings Analyser (its section 5) and ART Flow Steering (its section 6). For the second,
also check whether a custom skill can be attached to a Studio agent.
Pending, asked of AL (2026-11-07), not started: a lighter JSON object. The extraction of the fake Jira (1 244 issues) takes
24 s, of which the parse of the JSON is about 15 s (5 s for the search, 10 s for the changelogs): each JSON object is a
`Scripting.Dictionary`, and the first call on a fresh one costs about 0.65 ms on the dev machine (a `Collection` costs
nothing). Replacing it by a `Collection` with keys would save about 10 to 12 s, but it touches every reader of a parsed
object (`.Exists`, `.Keys`, `Item`: `modJson`, `modJira`, `modJiraConfig`, `modExtract`, `modText`, the `T_` tests), so
ask AL before starting. Measure first with the log and a scratch benchmark module (see the VBA speed traps).
Also pending, offered to AL: when the macro loses its state (Esc then End during a long step), the Dashboard stays on its
last status and Cancel does nothing; `AbortExtraction` and a slice that finds `mBusy` False could write a log line and
"Stopped" in `A3`.
