# jira2dashboard

jira2dashboard computes the six SAFe flow metrics (Velocity, Flow Time, Flow Load, Flow Efficiency, Flow
Distribution, Flow Predictability) and 29 data-quality checks from Jira Cloud, and shows them on a dashboard in an
Excel workbook. It is built for the RTE of an Agile Release Train. The workbook runs on the user's machine with
Excel alone: VBA reads Jira and Excel formulas compute every figure. There is no server, Python is not needed to
use it, and the data stays on the machine: the only network calls go to Jira, through the proxy.

## Contents

1. [Quick start](#1-quick-start)
2. [Overview](#2-overview)
3. [How it works](#3-how-it-works)
4. [Repository layout](#4-repository-layout)
5. [Requirements](#5-requirements)
6. [Scripts and commands](#6-scripts-and-commands)
7. [Typical workflows](#7-typical-workflows)
8. [Configuration](#8-configuration)
9. [Using the workbook](#9-using-the-workbook)
10. [Python extraction without VBA](#10-python-extraction-without-vba)
11. [Metrics and data checks](#11-metrics-and-data-checks)
12. [Logs and debugging](#12-logs-and-debugging)
13. [Testing](#13-testing)
14. [Release](#14-release)

## 1. Quick start

On a developer machine that meets the [requirements](#5-requirements), from the repository folder, in PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\setup.ps1 -Dev   # .venv with pywin32 and the test tools
.\scripts\run.ps1 check-machine                                   # 21 checks, all OK
.\scripts\run.ps1 test                                            # fast tests, all green, under a minute
.\scripts\run.ps1 build                                           # out\release\jira2dashboard.xlsm
.\scripts\manual-test.ps1                                         # fake Jira + filled copy, opened in a new Excel
```

In the copy that opens: **Enable Content**, click **Read the Jira configuration**, then **Extract from Jira and
compute**. The charts fill from the fake Jira (1200 issues). Stop the fake Jira with `.\scripts\stop-mock.ps1`.

Excel may crash right after the last search page (see [Symptoms and first steps](#symptoms-and-first-steps)): save
your other work first. The copy opens in an Excel of its own for that reason. `manual-test.ps1 -Python` tries the
path that does not use VBA.

## 2. Overview

**Deliverables**

| Deliverable | What it is | Made by |
|---|---|---|
| `jira2dashboard.xlsm` | The workbook to deliver. Needs Excel only. | `run.ps1 build` |
| `jira2dashboard.xlsx` | Fallback when the macros cannot run: values and SVG charts. | `run.ps1 extract` |
| `rovo/*.md` | Rovo Studio agents that read the published figures. | Written by hand |

**Constraints**

- Windows 11 and Excel for Microsoft 365 (dynamic array formulas: Excel 365 or 2021).
- Python is used to **build** and **test** the workbook. Its runtime needs pywin32 and the standard library only.
- The data read from Jira never leaves the machine where the workbook runs. Site values (URL, account, token,
  proxy, filter, field ids, as-of date, calendar) stay empty in this repository and are typed in the workbook.
- Mappings (statuses, transitions, fields, Flow Load mode) are **parameters in the workbook**, never code.

## 3. How it works

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
BUILD  developer machine, Python

  params/*.md ---+
  vba/*.bas  ----+---> run.ps1 build ---> jira2dashboard.xlsm
  flow/ plan ----+     build_workbook.py


RUN  user machine, Excel + VBA, no Python

  Jira Cloud REST API v3
        |
        |  1. read Jira       modHttp, modProxy, modJira (WinHTTP, proxy, Windows auth)
        |  2. write raw data  modExtract, modFlow
        v
  Issues, StatusIntervals, IssueLinks                       raw sheets (values)
        |
        |  3. rebuild         modFlowBuild, from the plan sheets
        v
  Controls, ItemFlow, DailyLoad, Metrics,                   computed sheets (formulas)
  FillRates, DataChecks, ChartData
        |
        v
  Dashboard: 8 charts, status cell A3


FALLBACK  Python, no VBA

  saved .xlsm ---> run.ps1 extract ---> standalone/ ---> Jira ---> flow/ ---> jira2dashboard.xlsx


TESTS

  mockjira/  fake Jira on 127.0.0.1, answers the workbook and standalone/
  flow/      Python reference, compared with the Excel formulas cell by cell
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

| Group | Sheets | Written by |
|---|---|---|
| Dashboard | `Dashboard` (buttons, status `A3`, charts, team `B2`, first steps `T1`), `Manual` | build |
| Parameters | `Config`, `Projects`, `Periods`, `PIs`, `Workflows`, `Transitions`, `FieldMapping` | build, then the user |
| Raw data | `Issues`, `StatusIntervals`, `IssueLinks` | VBA, each extraction |
| Results (formulas) | `Controls`, `ItemFlow`, `DailyLoad`, `Metrics`, `FillRates`, `DataChecks`, `ChartData` | VBA, from the plan |
| Plan (do not edit) | `FlowFormulas`, `ChartSpec`, `ChecksSpec` | build |
| Proposals from Jira | `JiraIssueTypes`, `JiraStatuses`, `JiraWorkflows`, `JiraTransitions`, `JiraFields` | VBA (`modJiraConfig`) |

### Key decisions

| Decision | Reason |
|---|---|
| VBA does the runtime work | The user's machine needs Excel only. |
| Every figure is a visible formula | Each number can be audited; the Python reference proves the formulas. |
| Formulas, charts and checks are data | One generic builder; a new metric or check needs no VBA change. |
| HTTP through WinHTTP (COM) | It answers a proxy's Windows authentication with the user's session. |
| Proxy API called on byte buffers | VBA user-defined Types crashed 64-bit Excel. |
| Fallback xlsx written with `zipfile` | The runtime stays pywin32 + standard library; no Excel needed. |

## 4. Repository layout

| Path | Content |
|---|---|
| `src/jira2dashboard/` | The Python package (table below) |
| `vba/` | 16 VBA modules; 13 are delivered (`modTests`, `modButtons`, `modBackground` are for tests only) |
| `params/` | The parameter tables written into the workbook (section [8](#8-configuration)) |
| `tests/` | `unit/` (no Excel), `features/` and `steps/` (Gherkin), `integration/` (Excel) |
| `scripts/` | The PowerShell entry points (section [6](#6-scripts-and-commands)) |
| `docs/` | Domain and technical references (frozen) |
| `rovo/` | Rovo Studio agents, their skills, the Confluence page template |
| `out/release/` | Delivered files: the built workbook, the release packages. Never cleaned. |
| `out/manual_test/` | The manual-test copy and its `-Python` result. Never cleaned. |
| `out/inter/` | Everything else: test workbooks, exports, test caches. Emptied by `build`. |
| `archive/` | Old code and probes, kept for reference |

**Inside `src/jira2dashboard/`**

| Module | Role |
|---|---|
| `build_workbook.py`, `clean.py` | Build the delivered workbook; empty `out\inter` |
| `flow/` | Python reference (metrics, checks) and the plan (formulas, charts) |
| `params/` | Read and validate `params/*.md`; write the parameter sheets |
| `standalone/` | Python extraction without VBA (`run.ps1 extract`) |
| `mockjira/`, `mock_workbook.py` | The fake Jira; the manual-test copy filled for it |
| `excel/`, `testbench/`, `cli/` | COM layer; VBA test runner; `check-machine` and the `--help` parser |
| `manual.py`, `paths.py` | Text of the sheet `Manual`; folder locations |

## 5. Requirements

**Developer machine**

- Windows 11, Windows PowerShell 5.1.
- Excel for Microsoft 365, with **File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust
  access to the VBA project object model** turned on (the build injects VBA). `check-machine` checks it.
- Python 3.14, started with the `py` launcher.

**User machine**

- To use the workbook: Excel for Windows (365 or 2021) with macros allowed. No Python.
- To build the workbook there, or to use the Python fallback: Python 3.14 and pywin32 (`setup.ps1`, without `-Dev`).

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
| `setup.ps1` | Creates `.venv` and installs the requirements | `-Dev`, `-Proxy` |
| `activate.ps1` | Activates `.venv` in the console (optional) | none |
| `run.ps1` | Runs a command of the project (table below) | `<command> [options]` |
| `start-mock.ps1` | Starts the fake Jira in a minimized window | `-Port`, `-Issues`, `-Hidden` |
| `stop-mock.ps1` | Stops the fake Jira that these scripts started | `-Port` |
| `manual-test.ps1` | Prepares and opens a manual try against the fake Jira | `-Workbook`, `-Port`, `-Python` |
| `release.ps1` | Builds a release; publishes it on GitHub with `-Publish` | `-Version`, `-Publish` |

- **`setup.ps1`** installs `requirements.txt` (pywin32), or `requirements-dev.txt` with `-Dev` (plus pytest,
  pytest-bdd, pytest-cov), then checks that pywin32 imports. `-Proxy http://host:port` is passed to pip.
- **`activate.ps1`** sets UTF-8 and `PYTHONPATH=src`, to run `python -m jira2dashboard...` by hand. `run.ps1`
  needs no activation.
- **`start-mock.ps1`** starts `run.ps1 mock-jira --anomalies` (default port 8099, 1200 issues) and waits until it
  answers. Account: `mock@example.com` / `mock-token`. It keeps the pid in `%TEMP%\jira2dashboard\` so that
  **`stop-mock.ps1`** kills only that process.
- **`manual-test.ps1`** stops an old fake Jira, builds the workbook if it is missing, writes the filled copy
  `out\manual_test\jira2dashboard4manualtest.xlsm`, starts the fake Jira and opens the copy in a new Excel. With
  `-Python` it runs `extract` on the copy instead and opens `out\manual_test\jira2dashboard.xlsx`.
- **`release.ps1`** needs a clean tree. It writes the code zip, the workbook and notes with SHA256 to
  `out\release\<version>`. `-Publish` creates the GitHub release: the files become public (section
  [14](#14-release)).

### Commands of `run.ps1`

| Command | What it does | Main options |
|---|---|---|
| `test` | Fast unit and Gherkin tests, no Excel | pytest arguments |
| `coverage` | `test` plus coverage; fails under 90 % | pytest arguments |
| `test-excel` | Scenarios that drive Excel (marker `excel`) | pytest arguments |
| `test-network` | Excel scenarios that also need the network | pytest arguments |
| `vba-tests` | Builds, compiles and runs the VBA test suite | `--offline`, `--bisect`, `--out` |
| `check-machine` | 21 checks of the machine | none |
| `mock-jira` | The fake Jira on `http://127.0.0.1:8099`, until Ctrl+C | `--anomalies`, `--port`, ... |
| `mock-workbook` | Writes the manual-test copy, filled for the fake Jira | `--workbook`, `--port` |
| `build` | Builds `out\release\jira2dashboard.xlsm` | `--out` |
| `clean` | Empties `out\inter`, except the test caches | none |
| `extract` | Reads Jira with Python, writes `jira2dashboard.xlsx` | `--workbook`, `--out`, `--team`, ... |

- `check-machine` checks Python, pywin32, Excel start and quit, access to the VBA project, compile, macros and child
  processes. Run it first on a new machine.
- `vba-tests --bisect` adds the modules one by one to find the first that does not compile. `--out name.xlsm`
  chooses another test workbook when `vba_tests.xlsm` is open.
- `mock-jira` also takes `--issues`, `--projects`, `--seed`, `--no-auth` and `--proxy-challenge`.
- `build --out name.xlsm` builds a variant in `out\inter`, not the workbook to deliver.
- `build` and `vba-tests` run `clean` first. `out\release` and `out\manual_test` are never cleaned.
- `extract` also takes `--log` and `--verbose` (section [10](#10-python-extraction-without-vba)).

## 7. Typical workflows

Run the commands of each workflow in this order.

| Workflow | Step | Command | Purpose |
|---|---|---|---|
| **First time on a developer machine** | 1 | `setup.ps1 -Dev` | The Python environment |
| | 2 | `run.ps1 check-machine` | Fix anything that is KO before going on |
| | 3 | `run.ps1 test` | The fast tests |
| | 4 | `run.ps1 build` | The workbook `out\release\jira2dashboard.xlsm` |
| | 5 | `manual-test.ps1` (or `-Python`) | Try it by hand against the fake Jira |
| | 6 | `stop-mock.ps1` | Stop the fake Jira after the try |
| **During a change** | 1 | `run.ps1 test` | At each step, under a minute |
| **Before finishing a change** | 1 | `run.ps1 coverage` | Fast tests and coverage of the pure code |
| | 2 | `run.ps1 vba-tests` | VBA tests, every coverage marker hit |
| | 3 | `run.ps1 test-excel` | Excel scenarios |
| | 4 | `run.ps1 test-network` | Excel and network scenarios |
| | 5 | `run.ps1 build` | The workbook to deliver |
| **Release** | 1 | `git commit`, `git push` | A clean tree; pushed before publishing |
| | 2 | `release.ps1 -Version v1.0.0` | Dry run: check `out\release\v1.0.0` |
| | 3 | `release.ps1 -Version v1.0.0 -Publish` | The GitHub release (public) |
| **User machine, `.xlsm` delivered alone** | 1 | none | Follow section [9](#9-using-the-workbook) |
| **User machine, code delivered** | 1 | `setup.ps1` | No `-Dev`; `-Proxy` if pip needs a proxy |
| | 2 | `run.ps1 check-machine` | The machine is ready |
| | 3 | `run.ps1 build` | Then fill the workbook and use its buttons |
| **Macros do not work** | 1 | `run.ps1 extract --workbook <workbook>` | Save the workbook first; needs `setup.ps1` (section [10](#10-python-extraction-without-vba)) |

## 8. Configuration

All parameters are sheets of the workbook. The files of `params/` are their sources in the repository: the build
validates them and writes them into the workbook. Change a value in the workbook, or in `params/` and rebuild. Never
in code. Each file of `params/` describes its columns and rules.

| Sheet | Source | Holds |
|---|---|---|
| `Config` | [`params/config.md`](params/config.md) | Connection, proxy, field ids, settings of the checks, log level |
| `Projects` | [`params/projects.md`](params/projects.md) | Jira projects to extract: one `Program`, the others `Team` |
| `Periods` | [`params/periods.md`](params/periods.md) | Iterations, half-open `[Start, End[`. Empty in the repo |
| `PIs` | [`params/pis.md`](params/pis.md) | Program Increment dates (check DC18). Empty in the repo |
| `Workflows` | [`params/workflows.md`](params/workflows.md) | Category of each status: Active, Wait, Done, Removed, Backlog |
| `Transitions` | [`params/transitions.md`](params/transitions.md) | Allowed moves (check DC28) |
| `FieldMapping` | [`params/field_mapping.md`](params/field_mapping.md) | One row per column of the ticket extract |

**Config columns:** Parameter, Value, Example, Mandatory, Description.

- **Value** is the only column read. **Example** shows the expected form.
- **Mandatory** = `True`: the button **Extract from Jira and compute** stops before any call to Jira and names the
  empty values (error 608). **Read the Jira configuration** needs only `BaseUrl`, `Email` and `ApiToken`.
- `False`: an empty value is allowed. A default is used, nothing is added (no filter, no proxy), or the checks that
  need it show "not computable".

**Main Config parameters** (full list with descriptions in [`params/config.md`](params/config.md)):

- **Connection:** `BaseUrl` (https only), `Email`, `ApiToken`.
- **Proxy:** `ProxyServer`, `ProxyBypass`, `ProxyWindowsAuth`. `ProxyServer` empty = the proxy of the Windows
  settings (WPAD, then the PAC script read from the network as a browser does, then a fixed proxy); `DIRECT` = no
  proxy; otherwise `host:port`. `ProxyWindowsAuth` = `Yes` answers the proxy's Windows authentication with the
  user's session; `No` never sends it.
- **Requests:** `TimeoutSeconds` (60), `MaxRetries` (4), `Filter` (extra JQL, no `ORDER BY`).
- **Computation:** `AsOf` (UTC date and time of the extraction), `FlowLoadMode` (`A` or `B`), `WorkTypeField`.
- **Data checks:** `FeatureTypes`, `SubTaskTypes`, `ParentRequiredTypes`, `PointsRequiredTypes`, `BlocksLinkType`,
  `FillRateThreshold` (90), and the field ids `StoryPointsField`, `BusinessValueField`, `PIField`, `TeamField`,
  `SprintField`, `TargetStartField`, `TargetEndField`.
- **Log:** `LogLevel` (`Steps` or `Errors`).

`BaseUrl`, `Email`, `ApiToken`, `ProxyServer`, `ProxyBypass`, `Filter`, `WorkTypeField`, `AsOf`, the field ids and
the calendars stay empty in the repository. A test refuses a token in `params/config.md`.

## 9. Using the workbook

The workbook repeats these steps on the `Dashboard` from cell `T1`, and in full in the sheet `Manual`.

**Delivery.** Either the file `jira2dashboard.xlsm` alone, or the code, built on the user's machine with
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
| Flow Time | Median and 85th percentile of the time from start to done |
| Flow Load | Issues in process at the end of each UTC day, averaged over the period |
| Flow Efficiency | Active time / (active + wait time) |
| Flow Distribution | Share of each work type among the issues done |
| Flow Predictability | Always "not computable": Jira holds no PI objectives |

- **Flow Time** leaves out the issues done without a valid start, or reopened; they are counted apart.
- **Flow Load** per day is in the sheet `DailyLoad`. `FlowLoadMode` A: an issue counts from its first start until
  done or removed, except while in a Backlog status. B: only from its last start (after the last return to backlog).
- **Flow Efficiency** is not computable while a status is Unclassified (check DC08).
- **Flow Distribution** reads `WorkTypeField`: Feature, Enabler, Defect, Autre; anything else is Unmapped.
- A figure that cannot be computed shows **"not computable"**, never 0. On a chart it is a gap with a note. The
  80-100 % band on Flow Predictability is a proposal (SAFe reference).

**Data checks.** `DataChecks` has one row per check DC01–DC29: the number of issues at fault (`Items`) and `OK`,
`KO` or `not computable`. To see the issues, filter `ItemFlow` on the column `DCnn` = 1. A check that needs a Jira
field not configured in `Config` is "not computable" and flags nobody.

| Area | Checks |
|---|---|
| Hierarchy and types | DC01–DC03, DC24 |
| Status history and workflow | DC04–DC10, DC27, DC28 |
| Sprints and story points | DC11, DC12, DC15, DC16 |
| Team and PI | DC13, DC14, DC17, DC18, DC29 |
| Dates, business value, fix versions | DC19–DC21, DC26 |
| Dependencies (sheet `IssueLinks`) | DC22, DC23 |
| Field fill rates (sheet `FillRates`) | DC25 |

The list with titles and severities is `CHECKS` in `src/jira2dashboard/flow/checks.py`. The interpretations are
proposals to validate with the teams.

## 12. Logs and debugging

### Where to look

| Name | Path | Written by |
|---|---|---|
| VBA run log | `%LOCALAPPDATA%\jira2dashboard\logs\jira2dashboard.log` | the workbook (`modLog`) |
| Python log | `%LOCALAPPDATA%\jira2dashboard\logs\python.log` | `run.ps1 extract` |
| Status cell, progress bar | `Dashboard!A3`, next to the buttons | the workbook |
| Sheets `Run` and `Errors` | in `jira2dashboard.xlsx` | `run.ps1 extract` |
| Unfinished-run flag | `HKCU\Software\VB and VBA Program Settings\jira2dashboard` | the workbook |
| VBA test log | `%TEMP%\jira2dashboard_vba_tests.log` | the VBA tests |
| Fake Jira pid file | `%TEMP%\jira2dashboard\mock-jira-<port>.pid` | `start-mock.ps1` |
| Fake Jira window | minimized PowerShell window | `start-mock.ps1` |
| Test workbooks | `out\inter\vba_tests.xlsm`, `out\inter\bisect_<n>.xlsm` | `vba-tests` |
| Manual-test copy | `out\manual_test\jira2dashboard4manualtest.xlsm` | `manual-test.ps1` |
| Chart exports | `out\inter\dashboard\*.png` | `flow_excel.feature` |
| Test caches | `out\inter\.coverage`, `out\inter\.pytest_cache` | pytest |
| Windows event log | Event Viewer > Windows Logs > Application, event 1000 | Windows |

What each one tells you:

- **VBA run log.** One line per step of each run: `RUN` at start (Excel version, workbook) and at end (`end ok`,
  `end cancelled`); `BEGIN`/`END` around each step (search page, changelog chunk, raw sheets, flow sheets); `ERROR`
  with the error number. The last `BEGIN` without `END` is the step where Excel stopped. `LogLevel` = `Errors`
  keeps only start, end and errors. Above 1 MB the old file becomes `jira2dashboard.1.log`.
- **Python log.** Each step, each Jira call that failed (status, call, attempt, wait, Jira's message), each error
  with its Python traceback. `--verbose` adds every call that worked; `--log` writes elsewhere. Above 1 MB the old
  file becomes `python.log.1`.
- **Status cell.** What the run is doing now; on failure, the message with its error number.
- **Sheets `Run` and `Errors`.** Counts and log file of the Python run; its warnings (retries, statuses not in
  `Workflows`).
- **Unfinished-run flag.** Set while a run is in progress. If Excel closed or crashed, the next run writes
  `RUN previous run did not end (started ...)`.
- **VBA test log.** The VBA tests never write to the real log.
- **Fake Jira pid file and window.** The pid and start time, so that `stop-mock.ps1` kills only what it started.
  The window shows the address, the account and the field ids; requests are not printed.
- **Test workbooks.** The workbook the VBA tests ran in; with `--bisect`, one per step.
- **Manual-test copy.** The copy filled for the fake Jira, and next to it `jira2dashboard.xlsx`, the result of
  `-Python`. Never cleaned.
- **Chart exports.** The Dashboard charts as images, to look at after an Excel test run.
- **Test caches.** Coverage data and the last failed tests (`--lf`).
- **Windows event log.** Event 1000 "Application Error" names the faulting module when Excel crashes (exit code
  `0xC0000005`).

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
| 602 | HTTP error after the retries, or `BaseUrl` not https | Read the status and the hint (below). |
| 603 | `Filter` contains `ORDER BY` | Remove it: the tool sets the order. |
| 604 | `AsOf` is empty or not a date | Type a UTC date and time: `2026-11-03 18:00`. |
| 505, 606 | Jira keeps sending the same page | Retry later. |
| 607 | Jira answered 200 with something that is not JSON | Usually a proxy or login page: check the proxy. |
| 608 | Mandatory values of `Config` are empty | Fill the values it names. |
| 610 | Cancelled by the user | None. |
| 611 | An extraction is already running | Wait for it, or cancel it. |
| 513, 514 | JSON that cannot be read or written | Note the log lines; retry. |
| 620, 621 | Read the Jira configuration: no project, or no end | Fill `Projects`; retry. |
| 701, 703, 704 | `Periods`: bad dates, no period, or end before start | Fix `Periods`. |
| 702, 709 | `Issues` is empty or has no column `Type` | Run the extraction again. |
| 705, 706 | A sheet is missing, or a column header differs | Rebuild the workbook (`run.ps1 build`). |
| 707 | Excel refused a formula of the plan | Note the table and column it names: a bug. |
| 708 | Excel too old: dynamic array formulas needed | Use Excel 365 or 2021. |

Hints of error 602: 401 = account or token (`Email`, `ApiToken`); 407 = proxy authentication (`ProxyServer`,
`ProxyWindowsAuth`); no answer = `BaseUrl`, `ProxyServer` or `ProxyBypass`.

### Symptoms and first steps

- **An extraction fails.** Read `A3`, then the `ERROR` line of the VBA log, then the table above.
- **Excel crashes during an extraction.** The last `BEGIN` without `END` gives the step. Right after
  `END search page N` it is the known open defect: a crash (`0xC0000005`) inside the VBA engine, more frequent after
  several extractions in the same Excel (analysis: `claude-code-handoff.md`, section 8b). Run again in a fresh
  Excel, or use `run.ps1 extract`. The Windows event log names the faulting module.
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
| Unit and pure scenarios | `tests/unit/`, `tests/features/` | `test`, `coverage` | no |
| Excel scenarios | marker `excel` | `test-excel` | yes, a fresh one per scenario |
| Excel + network scenarios | markers `excel` and `network` | `test-network` | yes |
| VBA unit tests | `vba/modTests.bas`: `T_` offline, `N_` network | `vba-tests` | yes |
| VBA against the fake Jira | `M_` tests, `@excel` scenario of `mock_jira.feature` | `test-excel` | yes |

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
