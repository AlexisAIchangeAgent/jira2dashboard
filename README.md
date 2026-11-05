# jira2dashboard

jira2dashboard computes the six SAFe flow metrics (Velocity, Flow Time, Flow Load, Flow Efficiency, Flow
Distribution, Flow Predictability) and 29 data-quality checks from Jira Cloud, and shows them on a dashboard in an
Excel workbook. It is built for the RTE of an Agile Release Train. The workbook runs on the client machine with Excel
alone: VBA reads Jira and Excel formulas compute every figure. There is no server, Python is not needed to use it,
and the data stays on the machine: the only network calls go to the client's Jira, through its proxy.

## Contents

1. [Overview](#1-overview)
2. [How it works](#2-how-it-works)
3. [Repository layout](#3-repository-layout)
4. [Requirements](#4-requirements)
5. [Quick start](#5-quick-start)
6. [Scripts and commands](#6-scripts-and-commands)
7. [Typical workflows](#7-typical-workflows)
8. [Configuration](#8-configuration)
9. [Using the workbook on the client machine](#9-using-the-workbook-on-the-client-machine)
10. [Python extraction without VBA](#10-python-extraction-without-vba)
11. [Metrics and data checks](#11-metrics-and-data-checks)
12. [Logs and debugging](#12-logs-and-debugging)
13. [Testing](#13-testing)
14. [Release](#14-release)
15. [Security and data handling](#15-security-and-data-handling)
16. [Known issues](#16-known-issues)
17. [Contributing](#17-contributing)
18. [Related documents](#18-related-documents)

## 1. Overview

**Deliverables**

| Deliverable | What it is | Made by |
|---|---|---|
| `jira2dashboard.xlsm` | The workbook to deliver: parameter sheets, VBA, buttons, dashboard. Needs Excel only. | `run.ps1 build` |
| `jira2dashboard.xlsx` | Fallback when the macros cannot run: the same data and results as values, charts as SVG pictures. | `run.ps1 extract` |
| `rovo/*.md` | Master copies of the Rovo Studio agents that read the published figures. Configuration, not code. | Written by hand |

**Constraints**

- Windows 11 and Excel for Microsoft 365 (dynamic array formulas: Excel 365 or 2021).
- Python is used to **build** and **test** the workbook. Its runtime needs pywin32 and the standard library only.
- Nothing from the client may leave the client machine: no data, no script. Client values (URL, account, token,
  proxy, filter, field ids, as-of date, calendar) stay empty in this repository and are typed in the workbook.
- Mappings (statuses, transitions, fields, Flow Load mode) are **parameters in the workbook**, never code.

## 2. How it works

### Principles

- **Python generates, Excel computes.** `build_workbook.py` writes the parameter sheets, the empty raw sheets and a
  *plan*, then injects and compiles the VBA modules. VBA writes only raw data. Every figure is an Excel formula that
  the user can read.
- **Formulas and charts are data.** `flow/formulas.py` and `flow/dashboard_spec.py` hold the formula text and the
  chart descriptions. They are written into the sheets `FlowFormulas`, `ChartSpec` and `ChecksSpec`. After each
  extraction, VBA (`modFlowBuild`) rebuilds every table, formula and chart from those sheets, sized to the new data.
  A metric or a check is added by editing the plan, not the VBA.
- **Parameters live in the workbook.** `params/*.md` are validated by `src/jira2dashboard/params/`, written into
  sheets, and read by VBA (`modParams`). A guard test refuses status, team or column names as literals in VBA.
- **A Python reference is the oracle.** `flow/` is a pure Python implementation of the same metrics and checks.
  The Excel scenarios run the delivered workbook against a fake Jira and compare Excel with it, cell by cell.

### Architecture

```
                 BUILD (developer machine)                             RUN (client machine)

 params/*.md --+
 vba/*.bas  ---+--> run.ps1 build --> jira2dashboard.xlsm --> Excel + VBA (no Python)
 flow/ plan ---+    build_workbook.py                          |
                                                               | 1. read Jira (WinHTTP, proxy, Windows auth)
       Jira Cloud REST API v3  <-------------------------------+    modHttp, modProxy, modJira
                                                               |
                                                               | 2. write the raw sheets (values)
                                                               |    modExtract, modFlow:
                                                               |    Issues, StatusIntervals, IssueLinks
                                                               |
                                                               | 3. rebuild from the plan (formulas)
                                                               |    modFlowBuild: Controls, ItemFlow, DailyLoad,
                                                               |    Metrics, FillRates, DataChecks, ChartData
                                                               v
                                                     Dashboard: 8 charts, status cell A3

 FALLBACK, no VBA:  saved .xlsm --> run.ps1 extract --> standalone/ --> Jira --> flow/ --> jira2dashboard.xlsx

 TESTS:  mockjira/ (fake Jira on 127.0.0.1) answers the workbook and standalone/;
         flow/ (Python reference) is compared with the Excel formulas, cell by cell.
```

### One extraction, step by step

1. The button **Extract from Jira and compute** (`modActions.ExtractAndCompute`) checks the mandatory values of
   `Config` (error 608) before any call to Jira.
2. The work runs in short slices scheduled with `Application.OnTime` (`modActions`, `modExtract`). Excel stays
   usable between slices, a progress bar moves, and **Cancel** stops the run.
3. `modJira` searches the projects of `Projects` (plus `Filter`) page by page, then fetches the changelogs in bulk.
   `modHttp` retries after 429 and 5xx (2 s × 2^n, at most 30 s, or `Retry-After` up to 120 s).
4. Only after Jira has answered, the previous results are cleared. `modExtract` and `modFlow` write `Issues`,
   `StatusIntervals` (one row per status interval) and `IssueLinks`.
5. `modFlowBuild.BuildFlowSheetsIn` rebuilds the computed sheets and the charts from the plan.
6. `modLog` writes each step to the local log (section [12](#12-logs-and-debugging)).

### Workbook sheets

| Sheets | Role | Written by |
|---|---|---|
| `Dashboard` | Buttons, status cell `A3`, progress bar, the 8 charts, team selector `B2`, first steps from `T1` | build, then VBA |
| `Manual` | The user manual | build (`manual.py`) |
| `Config`, `Projects`, `Periods`, `PIs`, `Workflows`, `Transitions`, `FieldMapping` | Parameters | build, then the user |
| `Issues`, `StatusIntervals`, `IssueLinks` | Raw data from Jira (values, plus helper formula columns) | VBA at each extraction |
| `Controls`, `ItemFlow`, `DailyLoad`, `Metrics`, `FillRates`, `DataChecks`, `ChartData` | Computed tables (formulas) | VBA, from the plan |
| `FlowFormulas`, `ChartSpec`, `ChecksSpec` | The plan: formulas, charts, checks. Do not edit. | build |
| `JiraIssueTypes`, `JiraStatuses`, `JiraWorkflows`, `JiraTransitions`, `JiraFields` | Proposals read from Jira | VBA (`modJiraConfig`) |

### Key decisions

| Decision | Reason |
|---|---|
| VBA does the runtime work | The client machine needs Excel only. |
| Every figure is a visible Excel formula | Each number can be audited in the workbook; the Python reference proves the formulas. |
| Formulas, charts and checks are data (the plan) | One generic builder in VBA; a new metric or check needs no VBA change. |
| HTTP through the WinHTTP COM object | It answers a proxy's Windows authentication (NTLM/Kerberos) with the user's session. |
| Proxy found with WinHTTP API calls on byte buffers | VBA user-defined Types crashed 64-bit Excel. |
| Python fallback writes the xlsx with `zipfile`, charts as SVG | The runtime stays pywin32 + standard library; no Excel is needed to write it. |

## 3. Repository layout

```
src/jira2dashboard/
  build_workbook.py     builds the workbook to deliver: out\release\jira2dashboard.xlsm (13 modules, no test module)
  clean.py              empties out\inter before a build (keeps the test caches); never touches out\release
  manual.py             text of the sheet Manual and of the first steps on the Dashboard
  mock_workbook.py      manual-test copy filled for the fake Jira: out\inter\jira2dashboard_mock.xlsm
  paths.py              locations: repo, vba, params, out\release, out\inter, children
  cli/                  args.py (shared --help parser), check_machine.py (21 checks of the machine)
  excel/com.py          COM layer (Excel, VBA project, workbook); pure helpers tested without Excel
  params/               mdtable.py (reads md tables), model.py (loads and validates), write.py (writes sheets)
  flow/                 Python reference and plan:
                          model.py (domain vocabulary), intervals.py, itemflow.py, load.py (Flow Load A/B),
                          metrics.py, checks.py (DC01-DC29), check_formulas.py, refs.py,
                          formulas.py (formula text), dashboard_spec.py (charts), blueprint.py (writes the plan),
                          sheets.py (COM: writes the plan and the manual, reads tables, exports charts)
  standalone/           Python path without VBA (run.ps1 extract): xlsx_read, params_book, proxy, http_winhttp,
                          jira_api, mapping, results, charts_svg, xlsx_write, __main__
  mockjira/             fake Jira Cloud: server.py (http.server, proxy, fault injection), api.py, jql.py,
                          dataset.py (hand-written cases), generator.py (seeded volume), anomalies.py (known
                          data-quality cases), config.py (configuration calls), flow_adapter.py, errors.py
  testbench/            builds, compiles and runs the VBA tests; children/ = scripts run as separate processes
vba/                    16 modules .bas (ASCII, LF). Delivered (13): modHttp, modProxy, modJira, modJiraConfig,
                        modJson, modText, modParams, modExtract, modFlow, modFlowBuild, modActions, modLog,
                        modCoverage. Test workbook only: modTests (the VBA tests), modButtons (its buttons),
                        modBackground (an older import of JSON files written outside Excel)
params/                 config.md, projects.md, periods.md, pis.md, workflows.md, transitions.md, field_mapping.md
tests/unit/             pure Python, no Excel (fast)
tests/features/         Gherkin scenarios (*.feature, French keywords)
tests/steps/            pytest-bdd steps
tests/integration/      tests marked excel (test_com_excel.py)
scripts/                setup, activate, run, start-mock, stop-mock, manual-test, release (*.ps1)
docs/                   7 context and reference documents (frozen)
rovo/                   Rovo Studio agents, skills and the Confluence page template
out/release/            what is delivered: the workbook of build and the release packages (never cleaned)
out/inter/              everything else: test workbooks, manual-test copy, exports, test caches (emptied by build)
archive/                old code and probes, kept for reference
```

## 4. Requirements

**Developer machine**

- Windows 11, Windows PowerShell 5.1.
- Excel for Microsoft 365, with **File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust
  access to the VBA project object model** turned on (the build injects VBA). `check-machine` checks it.
- Python 3.14, started with the `py` launcher.

**Client machine**

- To use the workbook: Excel for Windows (365 or 2021) with macros allowed. No Python.
- To build the workbook there, or to use the Python fallback: Python 3.14 and pywin32 (`setup.ps1`, without `-Dev`).

## 5. Quick start

From the repository folder, in PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\setup.ps1 -Dev   # .venv with pywin32 and the test tools
.\scripts\run.ps1 check-machine                                   # 21 checks, all OK
.\scripts\run.ps1 test                                            # fast tests, all green, under a minute
.\scripts\run.ps1 build                                           # out\release\jira2dashboard.xlsm
.\scripts\manual-test.ps1                                         # fake Jira + filled copy, opened in a new Excel
```

In the copy that opens: **Enable Content**, click **Read the Jira configuration**, then **Extract from Jira and
compute**. The charts fill from the fake Jira (1200 issues). Stop the fake Jira with `.\scripts\stop-mock.ps1`.

Excel may crash right after the last search page (section [16](#16-known-issues)): save your other work first. The
copy opens in an Excel of its own for that reason. `manual-test.ps1 -Python` tries the path that does not use VBA.

## 6. Scripts and commands

### Running a script

```powershell
.\scripts\run.ps1 test                                            # the usual form
powershell -ExecutionPolicy Bypass -File scripts\setup.ps1 -Dev   # when PowerShell refuses to run scripts
. .\scripts\activate.ps1                                          # activate.ps1 only: dot-sourced
.\scripts\manual-test.ps1 --help                                  # help of any script: --help, -Help or -?
.\scripts\run.ps1 build --help                                    # help of a run.ps1 command
```

- Script options start with `-` (`-Port 8100`). Options of a `run.ps1` command start with `--` (`--port 8100`)
  and reach Python or pytest unchanged.
- `run.ps1`, `setup.ps1`, `release.ps1` and `manual-test.ps1` move to the repository folder first: a relative
  path given to them starts from there.
- Call `run.ps1` directly from PowerShell. Do not wrap it in `cmd /c` and do not add `2>&1`: it stops as soon as
  Python writes to the error stream.
- `-?` works in a PowerShell console, not through `powershell -File`. `--help` works everywhere.

### Scripts

| Script | What it does | Options |
|---|---|---|
| `setup.ps1` | Creates `.venv` (Python 3.14), installs `requirements.txt` (pywin32), or `requirements-dev.txt` with `-Dev` (plus pytest, pytest-bdd, pytest-cov). Checks that pywin32 imports. | `-Dev`, `-Proxy http://host:port` |
| `activate.ps1` | Optional. Activates `.venv` in the console (UTF-8, `PYTHONPATH=src`) to run `python -m jira2dashboard...` by hand. | none |
| `run.ps1` | Runs a command of the project with the Python of `.venv`. No activation needed. | `<command> [options]` |
| `start-mock.ps1` | Starts the fake Jira with the known anomalies in a minimized window and waits until it answers. Account `mock@example.com` / `mock-token`. | `-Port 8099`, `-Issues 1200`, `-Hidden` |
| `stop-mock.ps1` | Stops the fake Jira that `start-mock.ps1` or `manual-test.ps1` started on that port, and nothing else. | `-Port 8099` |
| `manual-test.ps1` | Manual try: stops an old fake Jira, builds the workbook if missing, writes a filled copy, starts the fake Jira, opens the copy in a new Excel. `-Python` runs `extract` instead and opens `jira2dashboard.xlsx`. | `-Workbook C:\x.xlsm`, `-Port 8099`, `-Python` |
| `release.ps1` | Builds a release in `out\release\<version>`: the code zip, the workbook, notes with SHA256. Needs a clean tree. `-Publish` also creates the GitHub release (public). | `-Version v1.0.0`, `-Publish` |

### Commands of `run.ps1`

| Command | What it does | Main options |
|---|---|---|
| `test` | Fast unit and Gherkin tests, no Excel | any pytest argument |
| `coverage` | The same, plus the coverage of the pure code; fails under 90 % | any pytest argument |
| `test-excel` | The scenarios that drive Excel (marker `excel`, without `network`) | any pytest argument |
| `test-network` | The Excel scenarios that also need the network (markers `excel` and `network`) | any pytest argument |
| `vba-tests` | Builds, compiles and runs the VBA test suite; reports coverage of the VBA markers | `--offline`, `--bisect`, `--out name.xlsm` |
| `check-machine` | 21 checks of the machine: Python, pywin32, Excel, VBA project access, compile, child processes | none |
| `mock-jira` | The fake Jira on `http://127.0.0.1:8099`, until Ctrl+C | `--anomalies`, `--issues`, `--port`, `--projects`, `--seed`, `--no-auth`, `--proxy-challenge` |
| `mock-workbook` | A copy of a workbook with `Config`, `Periods` and `PIs` filled for the fake Jira | `--workbook`, `--port` |
| `build` | The workbook to deliver: `out\release\jira2dashboard.xlsm` | `--out name.xlsm` (goes to `out\inter`) |
| `clean` | Empties `out\inter` except the test caches (`build` and `vba-tests` do it first) | none |
| `extract` | Without VBA: reads Jira with Python and writes `jira2dashboard.xlsx` | `--workbook`, `--out`, `--team`, `--log`, `--verbose` |

## 7. Typical workflows

**First time on a developer machine**

1. `setup.ps1 -Dev`: the Python environment.
2. `run.ps1 check-machine`: fix anything that is KO before going on.
3. `run.ps1 test`: the fast tests.
4. `run.ps1 build`: the workbook `out\release\jira2dashboard.xlsm`.
5. `manual-test.ps1` (or `manual-test.ps1 -Python`): try it by hand against the fake Jira.
6. `stop-mock.ps1`: stop the fake Jira after the try.

**During a change:** `run.ps1 test` at each step (under a minute).

**Before finishing a change** (the full pass, in this order): `run.ps1 coverage`, `run.ps1 vba-tests`,
`run.ps1 test-excel`, `run.ps1 test-network`, `run.ps1 build`.

**Release**

1. Commit everything, and push if you publish.
2. `release.ps1 -Version v1.0.0`: dry run, nothing leaves the machine. Check `out\release\v1.0.0`.
3. `release.ps1 -Version v1.0.0 -Publish`: the GitHub release.

**Client machine**

- The `.xlsm` is delivered alone: no script. Follow section [9](#9-using-the-workbook-on-the-client-machine).
- The code is delivered: `setup.ps1` (no `-Dev`; `-Proxy` if pip needs a proxy), `run.ps1 check-machine`, then
  `run.ps1 build`. Then fill the workbook and use its buttons.
- The macros do not work: save the filled workbook, then `run.ps1 extract --workbook <workbook>`
  (section [10](#10-python-extraction-without-vba)). `setup.ps1` must have run.

## 8. Configuration

All parameters are sheets of the workbook. The files of `params/` are their sources in the repository: the build
validates them and writes them into the workbook. Change a value in the workbook, or in `params/` and rebuild. Never
in code.

| Sheet | Source | Columns | Holds |
|---|---|---|---|
| `Config` | [`params/config.md`](params/config.md) | Parameter, Value, Example, Mandatory, Description | Connection and proxy, field ids, settings of the checks, Flow Load mode, as-of date, log level |
| `Projects` | [`params/projects.md`](params/projects.md) | Project, Role, Team | Jira projects to extract; one `Program` (the features), the others `Team` |
| `Periods` | [`params/periods.md`](params/periods.md) | Period, Start, End | Iterations to compute, half-open `[Start, End[`. Empty in the repo |
| `PIs` | [`params/pis.md`](params/pis.md) | PI, Start, End | Program Increment dates (data check DC18). Empty in the repo |
| `Workflows` | [`params/workflows.md`](params/workflows.md) | Workflow, Order, Status, Category, IsStart, IsEnd, IsFree | Category of each status: Active, Wait, Done, Removed, Backlog |
| `Transitions` | [`params/transitions.md`](params/transitions.md) | Workflow, From, To, Type | Allowed moves (data check DC28) |
| `FieldMapping` | [`params/field_mapping.md`](params/field_mapping.md) | Extract, Logical, Type, CanBeEmpty, UsedBy, Role, IfMissing, TypeConfirmed | One row per column of the ticket extract |

**Config columns**

- **Value** is the only column read. **Example** shows the expected form.
- **Mandatory** = `True`: the button **Extract from Jira and compute** stops before any call to Jira and names the
  empty values (error 608). **Read the Jira configuration** needs only `BaseUrl`, `Email` and `ApiToken`.
- `False`: an empty value is allowed. A default is used, nothing is added (no filter, no proxy), or the checks that
  need it show "not computable".

**Main Config parameters** (the full list with descriptions is in [`params/config.md`](params/config.md))

| Group | Parameters |
|---|---|
| Connection | `BaseUrl` (https only), `Email`, `ApiToken` |
| Proxy | `ProxyServer` (empty = the Windows settings: WPAD, PAC, then fixed proxy; `DIRECT` = none; or `host:port`), `ProxyBypass`, `ProxyWindowsAuth` |
| Requests | `TimeoutSeconds` (60), `MaxRetries` (4), `Filter` (extra JQL, no `ORDER BY`) |
| Computation | `AsOf` (UTC date and time of the extraction), `FlowLoadMode` (`A` or `B`), `WorkTypeField` |
| Data checks | `FeatureTypes`, `SubTaskTypes`, `ParentRequiredTypes`, `PointsRequiredTypes`, `StoryPointsField`, `BusinessValueField`, `PIField`, `TeamField`, `SprintField`, `TargetStartField`, `TargetEndField`, `BlocksLinkType`, `FillRateThreshold` (90) |
| Log | `LogLevel` (`Steps` or `Errors`) |

`BaseUrl`, `Email`, `ApiToken`, `ProxyServer`, `ProxyBypass`, `Filter`, `WorkTypeField`, `AsOf`, the field ids and
the calendars stay empty in the repository. A test refuses a token in `params/config.md`.

## 9. Using the workbook on the client machine

The workbook repeats these steps on the `Dashboard` from cell `T1`, and in full in the sheet `Manual`.

**Delivery.** Either the file `jira2dashboard.xlsm` alone, or the code, built on the client machine with
`setup.ps1` then `run.ps1 build`.

1. **Unblock the file** if it came from a download or an e-mail (Mark of the Web makes Excel block the macros):
   right-click > Properties > **Unblock**, or `Unblock-File .\jira2dashboard.xlsm`. Open it and click
   **Enable Content**.
2. **Fill `Config`**: every value whose Mandatory is `True` (`BaseUrl`, `Email`, `ApiToken`, `AsOf` are empty in a
   new workbook). **Fill `Projects`** (the Program and the team projects) and **`Periods`** (the iterations).
3. Click **Read the Jira configuration** (`modJiraConfig`). It reads Jira for the projects of `Projects` and writes
   proposal sheets, never the parameter sheets:
   - `JiraIssueTypes`: the workflow of each issue type of each project;
   - `JiraStatuses`: the statuses of each type, with their Jira category;
   - `JiraWorkflows`: same first columns as `Workflows`. To do → `Backlog`, done → `Done`, in progress left empty
     (Active or Wait to decide). `IsFree = 1` when every status can move to it. `Order = 1` for the creation status;
   - `JiraTransitions`: same first columns as `Transitions`, one row per allowed move;
   - `JiraFields`: the field ids for the `...Field` rows of `Config`.

   Still to decide by hand: `Order`, `IsStart`, `IsEnd`, Active or Wait, `Removed` instead of `Done`, the `Type` of
   a move. Copy what fits into `Workflows`, `Transitions` and `Config`. Without the right "View (read-only)
   workflow", the other sheets are still filled and the status cell says so.
4. Click **Extract from Jira and compute** and wait for the progress bar. The first part (reading Jira, size
   unknown) fills 30 % of the bar as an estimate. **Cancel** stops the run before the raw sheets are written: the
   previous results stay. The run can be repeated at will: the computed sheets and charts are rebuilt to the new
   size.
5. Read the charts on the `Dashboard` (cell `B2` chooses the team of Flow Distribution; `All teams` = all teams),
   the figures in `Metrics` and the data checks in `DataChecks`.
6. **Before sending the workbook to anyone, empty `ApiToken`** in `Config`: the token is stored in clear text.

Do not edit `FlowFormulas`, `ChartSpec` and `ChecksSpec`: they are the plan that rebuilds the computed sheets.

## 10. Python extraction without VBA

When the macros cannot run (Excel crash, macros blocked), Python does their work and writes `jira2dashboard.xlsx`:

```powershell
.\scripts\run.ps1 extract --workbook C:\Data\jira2dashboard.xlsm   # writes C:\Data\jira2dashboard.xlsx
.\scripts\run.ps1 extract --team Alpha --verbose                   # Flow Distribution of Alpha, every call logged
.\scripts\manual-test.ps1 -Python                                  # try it against the fake Jira
```

1. It reads the parameter sheets of the filled workbook as a file (zip), without Excel: **save the workbook
   first**. Columns are found by header name. The same mandatory values as the macros must be filled.
2. It reads Jira like the macros: same proxy (`ProxyServer`, or the Windows settings), same Windows session sent to
   the proxy, same JQL, paging, retries and error numbers.
3. It computes with the Python reference (`flow/`, the oracle of the workbook) and writes, next to the workbook read
   (or `--out`), only values, no formula, no VBA:
   - `Run`: start, end, as-of date, counts, version, log file;
   - `Charts`: the 8 Dashboard charts as SVG pictures;
   - `Metrics`, `DataChecks`, `FillRates`, `ItemFlow`, `DailyLoad`: the results, with the workbook's column names;
   - `Issues`, `StatusIntervals`, `IssueLinks`: the data;
   - `Errors`: the warnings of the run;
   - the parameter sheets used, with `ApiToken` masked.
4. If Jira or a parameter fails, nothing is written: the previous file stays. If the old file is open in Excel, it
   says "close jira2dashboard.xlsx first".

Known differences from the workbook: a team with no item has no `Metrics` row, and `DailyLoad` stops at `AsOf`.

## 11. Metrics and data checks

Definitions and decisions: [`docs/safe-flow-metrics-and-data-checks.md`](docs/safe-flow-metrics-and-data-checks.md).
Metrics are computed per period (sheet `Periods`) and per team. The features of the Program project count under the
team "ART".

| Metric | In short |
|---|---|
| Velocity | Issues done in the period |
| Flow Time | Median and 85th percentile of the time from start to done. Issues done without a valid start, or reopened, are left out and counted apart. |
| Flow Load | Issues in process at the end of each UTC day (sheet `DailyLoad`), averaged over the period. `FlowLoadMode` A: from the first start until done or removed, except while in a Backlog status. B: only from the last start (after the last return to backlog). |
| Flow Efficiency | Active time / (active + wait time); not computable while a status is Unclassified |
| Flow Distribution | Share of each work type (`WorkTypeField`: Feature, Enabler, Defect, Autre; anything else is Unmapped) |
| Flow Predictability | Always "not computable": Jira holds no PI objectives |

A figure that cannot be computed shows **"not computable"**, never 0. On a chart it is a gap with a note. The
80-100 % band on Flow Predictability is a proposal (SAFe reference).

**Data checks.** `DataChecks` has one row per check DC01–DC29: the number of issues at fault (`Items`) and `OK`,
`KO` or `not computable`. To see the issues, filter `ItemFlow` on the column `DCnn` = 1. A check that needs a Jira
field not configured in `Config` is "not computable" and flags nobody. The 29 checks cover:

- hierarchy and types: DC01–DC03, DC24;
- status history and workflow: DC04–DC10, DC27, DC28;
- sprints and story points: DC11, DC12, DC15, DC16;
- team and PI: DC13, DC14, DC17, DC18, DC29;
- dates, business value and fix versions: DC19–DC21, DC26;
- dependencies ("Blocks" links, sheet `IssueLinks`): DC22, DC23;
- field fill rates (sheet `FillRates`, threshold `FillRateThreshold`): DC25.

The list with titles and severities is `CHECKS` in `src/jira2dashboard/flow/checks.py`. The interpretations are
proposals to validate with the teams.

## 12. Logs and debugging

### Where to look

| Name | Path | Written by | Purpose |
|---|---|---|---|
| VBA run log | `%LOCALAPPDATA%\jira2dashboard\logs\jira2dashboard.log` (above 1 MB the old one becomes `jira2dashboard.1.log`) | the workbook (`modLog`) | One line per step of each run: `RUN` at start (Excel version, workbook) and end (`end ok`, `end cancelled`), `BEGIN`/`END` around each step (search page, changelog chunk, raw sheets, flow sheets), `ERROR` with the error number. The last `BEGIN` without `END` is the step where Excel stopped. `LogLevel` = `Errors` keeps only start, end and errors. |
| Python log | `%LOCALAPPDATA%\jira2dashboard\logs\python.log` (above 1 MB the old one becomes `python.log.1`) | `run.ps1 extract` | Each step, each Jira call that failed (status, call, attempt, wait, Jira's message), each error with its Python traceback. `--verbose` adds every call that worked; `--log` writes elsewhere. |
| Status cell and progress bar | `Dashboard!A3`, next to the buttons | the workbook | What the run is doing now; on failure, the message with its error number. |
| Sheets `Run` and `Errors` | in `jira2dashboard.xlsx` | `run.ps1 extract` | Summary of the Python run (counts, log file) and its warnings (retries, statuses not in `Workflows`). |
| Unfinished-run flag | `HKCU\Software\VB and VBA Program Settings\jira2dashboard` | the workbook | Set while a run is in progress. If Excel closed or crashed, the next run writes `RUN previous run did not end (started ...)`. |
| VBA test log | `%TEMP%\jira2dashboard_vba_tests.log` | the VBA tests | The tests never write to the real log. |
| Fake Jira pid file | `%TEMP%\jira2dashboard\mock-jira-<port>.pid` | `start-mock.ps1` | Pid and start time, so that `stop-mock.ps1` kills only what it started. |
| Fake Jira window | minimized PowerShell window | `start-mock.ps1` | Shows the address, the account and the field ids. Requests are not printed. |
| Test workbooks | `out\inter\vba_tests.xlsm`, `out\inter\bisect_<n>.xlsm` | `vba-tests` | The workbook the VBA tests ran in; with `--bisect`, one per step. |
| Manual-test copy | `out\inter\jira2dashboard_mock.xlsm`, `out\inter\jira2dashboard.xlsx` | `manual-test.ps1` | The filled copy, and the result of `-Python`. |
| Chart exports | `out\inter\dashboard\*.png` | `flow_excel.feature` | The Dashboard charts as images, to look at after an Excel test run. |
| Test caches | `out\inter\.coverage`, `out\inter\.pytest_cache` | pytest | Coverage data and the last failed tests (`--lf`). |
| Windows event log | Event Viewer > Windows Logs > Application, event 1000 "Application Error" | Windows | Names the faulting module when Excel crashes (exit code `0xC0000005`). |

**What the logs hold.** Both logs stay on the machine (`%LOCALAPPDATA%`, not OneDrive). The VBA log never holds the
token, the e-mail, the URL or Jira's answers (errors 602 and 607 log their number only). The Python log never holds
the token or the `Authorization` header, but it does hold the JQL (projects and `Filter`), the workbook path, the
proxy found and Jira's messages. **Reread a log before sharing it.**

### Error numbers

The workbook shows them as `vbObjectError+n` in the log and in the status cell. The Python extraction uses the same
numbers for the same causes.

| Number | Cause | What to do |
|---|---|---|
| 600 | No project to extract | Fill the sheet `Projects`. |
| 601 | The search did not end after 2000 pages | Narrow `Filter`; retry later. |
| 602 | Jira answered an HTTP error after the retries, or `BaseUrl` does not start with https:// | Read the HTTP status and the hint in the message: 401 = account or token, 407 = proxy authentication, no answer = `BaseUrl` or proxy. |
| 603 | `Filter` contains `ORDER BY` | Remove it: the order is set by the tool. |
| 604 | `AsOf` is empty or not a date | Type a UTC date and time, for example `2026-11-03 18:00`. |
| 505, 606 | Jira keeps sending the same page (changelog, search) | Retry later. |
| 607 | Jira answered 200 with something that is not JSON | Usually a proxy or login page: check `ProxyServer`. |
| 608 | Mandatory values of `Config` are empty | Fill the values it names. |
| 610 | Cancelled by the user | None. |
| 611 | An extraction is already running | Wait for it, or cancel it. |
| 513, 514 | JSON that cannot be read or written | Note the log lines; retry. |
| 620, 621 | Read the Jira configuration: no project, or the project search did not end | Fill `Projects`; retry. |
| 701, 703, 704 | `Periods`: dates not valid, no period, or a period ending before it starts | Fix `Periods`. |
| 702, 709 | `Issues` is empty or has no column `Type` | Run the extraction again. |
| 705, 706 | A sheet is missing, or a column header differs | Rebuild the workbook (`run.ps1 build`). |
| 707 | Excel refused a formula of the plan | Note the table and column it names; it is a bug. |
| 708 | Excel too old: dynamic array formulas needed | Use Excel 365 or 2021. |

### Symptoms and first steps

- **An extraction fails.** Read `A3`, then the `ERROR` line of the VBA log, then the table above.
- **Excel crashes during an extraction.** The last `BEGIN` without `END` gives the step. Right after
  `END search page N` it is the known defect (section [16](#16-known-issues)): run again in a fresh Excel, or use
  `run.ps1 extract`. The Windows event log names the faulting module.
- **Excel hangs.** A modal VBA error dialog is usually open behind the window: close it. Never kill an `EXCEL.EXE`
  the scripts did not start; the user may have work open in it.
- **Proxy problems.** The Python log has a line `END proxy: ...` saying what was found (WPAD, PAC, fixed, direct).
  Set `ProxyServer` to `host:port` or to `DIRECT` to skip detection.
- **The VBA does not compile.** `run.ps1 vba-tests --bisect` adds the modules one by one and names the first that
  fails.
- **The machine is not ready** (VBA project access, Excel start). `run.ps1 check-machine`.
- **Figures look wrong.** Check `DataChecks` first, then filter `ItemFlow` on a `DCnn` column = 1. A status in
  `Unclassified` (DC08) makes Flow Efficiency not computable.
- **An Excel test is flaky.** Watch the `EXCEL.EXE` exit code. `J2D_ONE_EXCEL=1` runs a whole module in one Excel
  (stress mode). A fix is proven by 5 green module runs in a row, and compared with the same runs of the last
  commit.
- **One test only:** `.\scripts\run.ps1 test tests/unit/test_checks.py::test_name`, or
  `.\scripts\run.ps1 test-excel tests/steps/test_flow_excel.py -k "contr" -q`.

## 13. Testing

**Approach.** BDD outside, TDD inside. A feature starts as a Gherkin scenario (`tests/features/*.feature`, French
keywords, steps in `tests/steps/`), then unit tests (`tests/unit/`) drive the code in red-green-refactor steps. A
feature is done when its scenario passes on the real workbook. A new test must discriminate: break the code on
purpose, watch the test fail, restore.

| Level | Where | Run with | Excel |
|---|---|---|---|
| Unit and pure scenarios | `tests/unit/`, `tests/features/` | `run.ps1 test` / `coverage` | no |
| Excel scenarios | marker `excel` | `run.ps1 test-excel` | yes, a fresh Excel per scenario |
| Excel + network scenarios | markers `excel` and `network` | `run.ps1 test-network` | yes |
| VBA unit tests | `vba/modTests.bas`: `T_` (offline), `N_` (network) | `run.ps1 vba-tests` | yes |
| VBA against the fake Jira | `M_` tests, from the `@excel` scenario of `mock_jira.feature` | `run.ps1 test-excel` | yes |

`test` and `coverage` leave out the markers `excel` and `network` (`addopts` in `pyproject.toml`).

**Oracle.** `tests/features/flow_excel.feature` runs the delivered workbook against the fake Jira and compares
`ItemFlow`, `DailyLoad`, `Metrics` and `DataChecks` with the Python reference cell by cell, in Flow Load modes A
and B. `standalone.feature` does the same for the Python extraction.

**Fake Jira** (`mockjira/`). A stdlib HTTP server: `search/jql` with token paging, changelog bulk fetch, the
configuration calls, fault injection (any status, `Retry-After`, a custom body), and an HTTP proxy (absolute-form
requests and `CONNECT`; `--proxy-challenge Negotiate` answers 407 until `Proxy-Authorization` is sent).
`--anomalies` adds the known data-quality cases that the data checks must find. It does not simulate TLS inside the
tunnel nor a real NTLM/Kerberos exchange.

**Coverage.**

- Pure Python ≥ 90 %, enforced by `run.ps1 coverage`. The code that needs Excel, the registry or a child process
  is listed as omitted in `pyproject.toml` and is exercised by the Excel tests, `check-machine` and `vba-tests`.
- VBA branches carry `Cov "ID"` markers. `vba-tests` requires 100 % of them hit, and a unit test pins their count.

**Before finishing a change:** `coverage`, `vba-tests`, `test-excel`, `test-network`, `build`.

## 14. Release

```powershell
.\scripts\release.ps1 -Version v1.0.0            # dry run: out\release\v1.0.0, nothing leaves the machine
.\scripts\release.ps1 -Version v1.0.0 -Publish   # also the GitHub release and its tag (public)
```

- The working tree must be clean. With `-Publish`, the commit must already be pushed and `gh` logged in.
- `out\release\<version>` holds `jira2dashboard-<version>.zip` (the code of HEAD, `git archive`), the workbook built
  from the same commit, and `notes.md` with the SHA256 of both.
- The zip leaves out the `export-ignore` paths of `.gitattributes`: `CLAUDE.md`, `claude-code-handoff.md`,
  `archive/`, `.claude/`, `.venv/`. `tests/unit/test_release_contents.py` pins the list. The "Source code" links
  that GitHub adds follow the same rules.
- `-Publish` makes the files public: decide it explicitly.

## 15. Security and data handling

- **Nothing leaves the client machine.** Requests go only to Jira at the `BaseUrl` of `Config`, through the proxy
  when one is set or found. With `ProxyServer` empty, proxy detection also reads the network's WPAD/PAC script, as
  a browser does. Jira's answers stay in the workbook; the logs stay in `%LOCALAPPDATA%`.
- **API token.** Typed in `Config` only, stored in clear text in the workbook: empty it before sharing the
  workbook. It is masked in `jira2dashboard.xlsx`, never written to any log, and a test refuses it in
  `params/config.md`.
- **Transport.** `BaseUrl` must be https (error 602 otherwise; `http` only for the local fake Jira). Basic
  authentication with e-mail and API token.
- **Proxy.** `ProxyWindowsAuth` = `Yes` answers the proxy's Windows authentication with the user's session, for
  the proxy typed in `ProxyServer` or found in the Windows settings. `No` never sends it.
- **Reading files.** The Python reader refuses an XML part with a DOCTYPE (no external entities, no entity
  expansion).
- **Logs** are designed to hold no secret, but the Python log holds the JQL and Jira's messages: reread before
  sharing (section [12](#12-logs-and-debugging)).
- **Excel processes.** The scripts and tests only ever close the Excel they started.

## 16. Known issues

- **Excel crash during the extraction** (`0xC0000005`, inside the VBA engine), right after the last search page
  (`END search page N` in the log), more often after several extractions in the same Excel. The Excel scenarios open
  a fresh Excel per scenario for this reason. Analysis: `claude-code-handoff.md`, section 8b. Workarounds: run the
  extraction in a fresh Excel, or use `run.ps1 extract`.
- **Not testable here, to check at the client:** a real PAC or WPAD network; the Windows session sent to a real
  proxy (the fake proxy accepts any `Proxy-Authorization`, and a POST through its `--proxy-challenge` mode arrives
  empty); the shapes of the Jira configuration API on a real Jira; the Rovo agents in Rovo Studio.
- **Pending feature:** a workflow per issue type (scenarios in `workflow_by_type.feature`, not bound yet).

## 17. Contributing

**Conventions**

- Code and docs in English. Domain vocabulary from `flow/model.py` (Issue, Interval, ItemFlow, Rules...).
- KISS and YAGNI. Logic in pure modules; COM and I/O in thin edges, testable without Excel.
- VBA: ASCII only, LF line endings, `Option Explicit`; button macros are public subs without arguments. No status,
  team or column name as a literal (guard test `tests/unit/test_parameters_not_in_code.py`).
- Every script and command answers `--help` (`tests/unit/test_help.py`). A new Python entry point parses its
  arguments first with `cli/args.parser`; a new `.ps1` gets comment-based help.
- Keep the three sides in step: the VBA, the Python reference `flow/`, and the Python path `standalone/`. A formula
  change needs the reference change, and an extraction rule change needs both extractions.
- When a `.bas` changes, update its MD5 in the table of `claude-code-handoff.md`. The 7 files of `docs/` are never
  edited.
- Commit messages end with a `Co-Authored-By` line when an assistant wrote the change.

**Adding a data check**

1. `flow/checks.py`: a `Check` row and the rule in `issue_checks` (Python reference).
2. `flow/formulas.py`: a `DCnn` column in `ITEMFLOW_COLUMNS` (helper columns in `INTERVAL_HELPERS` if needed).
3. If it needs a new Jira field: a raw column in `ISSUES`, `ExtractFields` and the row in `modExtract.bas`, the
   fixtures in `modTests.bas`, and the field in `mockjira/` (`FlowIssue` and `flow_adapter.py`).
4. `mockjira/anomalies.py`: a known case with the ids it must trigger; the data-check scenario of
   `flow_excel.feature` compares it with the workbook.
5. Update the counts (marker test, the Config row count in `params.feature` if a parameter is added), the MD5 table
   and the handoff.

**Adding a metric or a chart:** edit the plan (`flow/formulas.py`, `flow/dashboard_spec.py`) and the reference
(`flow/metrics.py`), not the VBA. Update `rovo/confluence_flow_metrics_page.md` when column names or chart titles
change.

## 18. Related documents

| Document | Content |
|---|---|
| [`docs/safe-flow-metrics-and-data-checks.md`](docs/safe-flow-metrics-and-data-checks.md) | Metric definitions, statuses, transitions, field mapping, data checks: the domain reference |
| [`docs/jira-cloud-api-vba-reference.md`](docs/jira-cloud-api-vba-reference.md) | Jira Cloud REST API as used from VBA |
| [`docs/pywin32-excel-vba-reference.md`](docs/pywin32-excel-vba-reference.md) | pywin32, Excel COM and VBA traps |
| [`docs/rovo-guide.md`](docs/rovo-guide.md) | Original Rovo guide (frozen; the current copy is `rovo/rovo-guide.md`) |
| `docs/cookbook_python.md`, `docs/vba-modules-tested.md` | Code copies from the start of the project, not maintained (the code lives in `src/` and `vba/`) |
| `docs/copilot-instructions.md` | Earlier instructions for GitHub Copilot, kept as context |
| [`params/`](params/) | The parameter tables and their rules |
| [`rovo/`](rovo/) | Rovo agents: Settings Analyser, Flow Metrics by Type, ART Flow Steering (skills in `rovo/skills/`), the Confluence page template, the Jira data-quality dashboard |
| `claude-code-handoff.md` | Roadmap, decisions, code review dispositions, MD5 of the `.bas` (not in the release zip) |
| `CLAUDE.md` | Guidance for Claude Code in this repository (not in the release zip) |
