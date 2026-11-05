# Handoff to Claude Code: xlsm generation with VBA, Jira extraction, SAFe flow metrics, Rovo

Date: 2026-11-03. Owner: AL (Inspearit, AI trainer and developer). Written by Claude (claude.ai chat + desktop link).

> **Update (repo restructured, step 1 done).** The product is now **jira2dashboard** and the folder is
> **`C:\dev\jira2dashboard`** (a Python repo, see `README.md`: layout, commands, TDD/BDD loops). Claude now
> writes the Python code; Copilot no longer generates anything; AL only tests the workbook at the client.
> Sections 1-3 and 5-9 below are the original context. **Sections 4, 6 and 10 are updated.**
> `docs\cookbook_python.md` and `docs\vba-modules-tested.md` are unmaintained copies of code (the code lives in
> `src\` and `vba\`). The old rule "`bootstrap_test` must print 15/15" no longer applies: the safety net is
> `scripts\run.ps1 test`, `test-excel`, `test-network`, `vba-tests` and `check-machine`.

## 1. Working with AL (read first)

- Answer in **French**, short and simple sentences. The **documentation files are in English**.
- AL wants to be **asked**, not handed filled-in deliverables. Mark every idea of your own as a **proposal**;
  AL validates line by line.
- When you mention a file, always give its **relative path**.
- AL is Jira admin at the client. Nothing from the client may leave the client machine (no scripts, no data).

## 2. Goal

AL must make an **offline Copilot (GPT-5.6, Copilot Chat) on a client machine** generate Python scripts
(pywin32) that build a macro-enabled workbook (`.xlsm`) with VBA inside. The workbook:
1. extracts all issues (all fields, custom field names) and full changelogs of a SAFe train from **Jira Cloud**
   (one Program project with features, several team projects with stories, tasks, sub-tasks; one project has
   1200 issues) through a **proxy with Windows auth**; Jira auth = email + API token typed in a Config sheet;
2. checks data quality (checks DC01-DC29), lists missing data for flow metrics and for **Jira Plans**;
3. computes the six **SAFe flow metrics** with every intermediate step visible (traceability);
4. shows tests in sheets (Input, Expected, Actual, OK/KO), buttons, progress bar; coverage >= 90 %.

Copilot gets everything as **Markdown files only** (Outlook blocks `.py` and `.bas`).

## 3. Deliverables (7 Markdown files, English)

They are in **`docs\`** in this folder. On 2026-11-03 18:05 their MD5 was checked: byte-identical to the
versions produced in the claude.ai session. They are the **source of truth** for the client delivery.

| File | MD5 |
|---|---|
| `docs\copilot-instructions.md` | 8ae72b7579a0020a484540dddc8bade3 |
| `docs\cookbook_python.md` | 206ef9e994ca949203b0e21f9b927b8b |
| `docs\vba-modules-tested.md` | ffca4f890ba43e1f4428c10e58a500b4 |
| `docs\pywin32-excel-vba-reference.md` | 168e5cc0ea443f22fff549144278b47c |
| `docs\jira-cloud-api-vba-reference.md` | 63afd27d42e964464b54a3eb3e5281eb |
| `docs\safe-flow-metrics-and-data-checks.md` | 013d541f1ea197aa859fe75234657cb2 |
| `docs\rovo-guide.md` | d89fdf6ec1a0feab5ce01d4a5817e431 |

If you edit one of them, the MD5 above changes: update this table.

| File | Content |
|---|---|
| `copilot-instructions.md` | Rules for Copilot: sources priority, 3 action tiers, questions to ask, project layout, function design, VBA rules, traps table (7.1), tests, coverage >= 90 % (pytest-cov or `trace`), answer format (one step per answer), self-check, definition of done, Jira workbook summary |
| `cookbook_python.md` | All 7 Python files with path + MD5: `tools/cookbook_extract.py` (bootstrap), `xlsm_ref.py`, `_check_child.py`, `_hang_child.py`, `check_reference.py`, `run_vba_tests.py`, `_run_tests_child.py` |
| `vba-modules-tested.md` | The 9 tested VBA modules, each under a heading with path `vba/<module>.bas` + MD5 |
| `pywin32-excel-vba-reference.md` | Tested pywin32 + Excel COM reference (session, trust check, sheets, data, dates, `put_en_us`, `as_cell`, VBA components, buttons, save, compile, child processes, delivery) |
| `jira-cloud-api-vba-reference.md` | Jira Cloud endpoints and limits, auth, proxy mapping from PowerShell, rate limits, options A/B, workbook layout, extraction flow, VBA traps, pagination schemes |
| `safe-flow-metrics-and-data-checks.md` | Workflows and Transitions sheets (all status categories decided), FieldMapping, flow metric formulas on `StatusIntervals` / `ItemFlow`, traceability rules, checks DC01-DC29, missing data report, Plans readiness |
| `rovo-guide.md` | Rovo facts, prompt pattern, prompts P00-P10, one agent with subagents, skills (beta) as SKILL.md, prevention (saved filters, workflow validators, comment-only automations, weekly agent summary), prompt testing, capitalisation, Excel alignment |

Bootstrap at the client: put the md files in `docs\`, copy-paste the first block of `cookbook_python.md`
section 4 into `tools\cookbook_extract.py`, run
`py -3.14 tools/cookbook_extract.py docs/cookbook_python.md docs/vba-modules-tested.md`.
Tested from an empty folder: 16 files rebuilt byte-identical; tampering gives `MD5 MISMATCH` and writes nothing.
Also tested **on this Windows machine** from `docs\` (`bootstrap_test.py`, output in `bootstrap_test\`):
`all files OK`, 15/15 files identical to the tested `.py` and `.bas` files of this folder.

**Rule:** no `C:\dev` (or any local path) inside the md files: this folder is AL's machine only.

## 4. This folder (`C:\dev\jira2dashboard`, AL's dev machine)

Created under `C:\dev` to avoid antivirus issues (renamed from `xlsm-check`). Holds the code and the test bench.

| Path | What |
|---|---|
| `docs\*.md` | The 7 context and parameter files (section 3) |
| `src\jira2dashboard\excel\com.py` | COM helpers (was `xlsm_ref.py`) |
| `src\jira2dashboard\testbench\run_vba_tests.py` | VBA harness (options `--offline`, `--bisect`, `--out name.xlsm`) |
| `src\jira2dashboard\testbench\children\` | Children: `_check_child.py` (compile + Ping), `_run_tests_child.py`, `_hang_child.py` (timeout test) |
| `src\jira2dashboard\cli\check_machine.py` | Machine readiness check, 21 checks (was `check_reference.py`) |
| `vba\*.bas` | The 10 tested modules (MD5 below); `modParams.bas` (E2) reads the parameter sheets |
| `params\*.md` | **Parameters of the application** (E2): `config.md`, `projects.md`, `workflows.md`, `transitions.md`, `field_mapping.md`, one table each. `src\jira2dashboard\params\` reads and validates them (`mdtable.py`, `model.py`) and writes them into the sheets `Config`, `Projects`, `Workflows`, `Transitions`, `FieldMapping` (`write.py`). **`Projects`** (`Project | Role | Team`) lists the projects to extract and says which one is the Program (exactly one, default `TTT`, the first key of the extraction script; the other keys are the script's anonymised keys, to replace in the workbook); `modParams.ProjectKeys` (Program first), `ProgramProject`, `TeamOfProject` read it. **`Config.Filter`** is an extra JQL restriction added with `AND` to the project list (`modJira.ExtractionJql`); it must not contain `ORDER BY`. Tests keep them equal to `docs\safe-flow-metrics-and-data-checks.md`. **`Config` holds the connection and the proxy** (`BaseUrl`, `Email`, `ApiToken`, `ProxyServer`, `ProxyBypass`, `ProxyWindowsAuth`, `TimeoutSeconds`, `MaxRetries`): the client values stay empty in the repo, `ApiToken` is refused if filled in `config.md`, and everything is typed in the workbook. `modJira.JiraConfigFromSheet` turns the sheet into the HTTP configuration |
| `src\jira2dashboard\flow\` | **Reference implementation of the six flow metrics** (E4 stage 1, pure Python, no Excel), written from `docs\safe-flow-metrics-and-data-checks.md` section 2 and used as the oracle for the Excel formulas. Decisions taken with AL on 2026-11-04: DC01-DC29 are a separate step after E4; Flow Time and Flow Efficiency leave out issues done without a valid start (SkippedStart) or reopened after Done (counted in `flow_time_excluded`, so included + excluded = done); Flow Load is a snapshot at the end of each UTC day on half-open intervals [entered, left), mode A by default (`Config.FlowLoadMode`), days without a finished snapshot at `Config.AsOf` are left out; Program features are reported under the team `ART`. Periods come from the sheet `Periods` (empty in the repo). A figure that cannot be computed is the text `not computable`, never 0; Flow Predictability is always `not computable` (no PI objectives in Jira). Tests: 9 hand-worked issues (`tests\unit\test_flow_reference.py`), properties on 1200 generated issues (`test_flow_properties.py`), Gherkin per metric (`tests\features\flow_metrics.feature`) |
| `src\jira2dashboard\flow\formulas.py`, `sheets.py` | **The Excel side of E4 (stage 2).** `formulas.py` holds the formulas as text (pure, unit-tested: every `Table[Column]` and `[@Column]` resolves, no status name or team name appears in a formula, same vocabulary as the Python reference). **Since E6 the tables and formulas are built by VBA (`modFlowBuild`) from the sheet `FlowFormulas`; `sheets.py` now only writes that plan and reads tables for tests.** Historically `sheets.py` (COM) turned the raw sheets `Issues` and `StatusIntervals` into tables, adds helper columns (IsStart, IsEnd, PrevCategory, Team, Bucket, FirstStart, LeftEff, Counted) and builds the tables `ItemFlow` (one row per issue, docs 2.1), `DailyLoad` (day x team) and `Metrics` (period x team, docs 2.2) with **visible formulas** (`MINIFS`, `MAXIFS`, `SUMIFS`, `COUNTIFS`, `FILTER`, `PERCENTILE.INC`...). The control cells `Metrics!B1` (as-of date, `AsOfUtc`) and `Metrics!B2` (mode, `cfg_FlowLoadMode`) are formulas that read `Config`; change `FlowLoadMode` there and everything recalculates. Dates are compared with a half-second margin (Excel turns a bound into text with 15 digits). **Differential test** `tests\features\flow_excel.feature`: fake Jira (250 issues) -> VBA extraction -> formulas -> cell-by-cell comparison with the Python reference, for `StatusIntervals`, `ItemFlow`, `DailyLoad` and `Metrics`, in mode A then B; a deliberate mutation of a formula was checked to make it fail. (Done in E6: the delivered workbook rebuilds everything itself.) **Excel traps found:** a formula typed in one cell of a table column becomes the formula of the whole column (write key columns before creating the table); a structured reference to a column whose name contains `_` is refused (`[@WipCount_Feature]`), so column names have none |
| `src\jira2dashboard\flow\dashboard_spec.py`, `sheets.py` (`export_charts`), `blueprint.py` | **Dashboard (E5, built by VBA since E6: `ChartSpec` sheet, `modFlowBuild.BuildChart`).** `dashboard_spec.py` (pure, unit-tested in `tests\unit\test_dashboard_spec.py`) lists the six charts (Velocity, Flow Time median, Flow Load average, Flow Efficiency, Flow Distribution, Flow Predictability), the formulas of the chart data and the layout. `build_dashboard` writes the sheet `ChartData` (one block per chart: series names, period labels, one formula per point, relative references written once per block) and the sheet `Dashboard` (native Excel charts, one per metric, **one series per team**; a note under each chart counts the points "not computable"). A "not computable" figure is `#N/A`: the chart shows a **gap**, never 0 (`DisplayBlanksAs = 1`; the series values read back as `None`). Flow Distribution has one series per work type and a team selector (`Dashboard!B2`, name `dashTeam`, list = `All teams` + the teams of `Projects`); "All teams" weights each team's share by its velocity. `export_charts` saves each chart as `out\dashboard\<name>.png`. Decisions with AL on 2026-11-04: one chart per metric with one series per team; not computable = gap + note under the chart. Test: the scenario of `flow_excel.feature` continues after the Metrics comparison (ChartData cell by cell against the Python reference, in mode A then B; the series of every chart read back from Excel; the selector; PNG files). **Traps:** `Chart.Export` writes an EMPTY file for a chart that is off the visible part of the sheet: activate the sheet and the chart object first; `Series.Values` returns `None` for `#N/A`. **Not done:** Flow Time P85 and Flow Load at period end are in `Metrics` but have no chart; no explicit colour palette (nine series use Excel's default colours, a few look alike); the tables and charts are still built by Python after the extraction (E6) |
| `src\jira2dashboard\mockjira\` | **Fake Jira Cloud** (E3), stdlib `http.server` on 127.0.0.1: `GET/POST /rest/api/3/search/jql` (token paging, `expand=names`), `POST /search/approximate-count`, `POST /changelog/bulkfetch`, `GET /field`, `GET /myself`, and the Jira configuration (`mockjira\config.py`, built from the parameters): `GET /project/search`, `GET /project/{key}/statuses`, `POST /workflows/preview`; `410` on the removed `/search`; faults (429, 500/503 with or without `Retry-After`) injected from Python (`fail_next`) or over HTTP (`POST /__mock__/fail`); Basic auth `base64("email:token")` checked when credentials are set (plain-text 401 like Jira); usable as an HTTP proxy: absolute-form requests and a plain `CONNECT` tunnel, with an optional proxy challenge (`proxy_challenge="Negotiate"`: 407 until the request carries `Proxy-Authorization`; any value is accepted, so a real NTLM/Kerberos token is not checked). Modelled on the existing PowerShell extraction script (project keys, BV/SP/PI custom field ids, Basic auth, `/rest/api/3` paths). JQL understood: `project in (...)` plus restrictions joined by `AND`: `created` / `updated` (>=, >, <=, < with `"2026-10-01"` or `-30d` / `-2w`, counted from the dataset's as-of date), `issuetype`, `status` (= or `in`), parentheses, `ORDER BY key|created|updated`; anything else is a 400 that says so. Not simulated: TLS 1.2 inside the tunnel and the real NTLM/Kerberos exchange (verify at the client). Data = issue status histories (`dataset.py`; `issue()` builds hand-written cases, `generator.py` builds a seeded volume), so search and changelog never disagree. Run it for Excel: `scripts\run.ps1 mock-jira` (`http://127.0.0.1:8099` by default), then put that URL in `BaseUrl` |
| `tests\` | `unit\` (no Excel), `features\` + `steps\` (Gherkin, pytest-bdd), `integration\` (Excel) |
| `scripts\` | `setup.ps1`, `activate.ps1`, `run.ps1` (PowerShell 5.1) |
| `src\jira2dashboard\build_workbook.py`, `vba\modFlowBuild.bas`, `modActions.bas` | **Delivery (E6).** `scripts\run.ps1 build` makes `out\jira2dashboard.xlsm` (12 modules with `modLog` and `modJiraConfig`, no test module). Button **Extract from Jira and compute** on the Dashboard runs `modActions.BtnExtractAndCompute`: extraction, then `BuildFlowSheetsIn`. Scenarios of `flow_excel.feature` run this macro on the delivered workbook against the fake Jira and compare with the Python reference, twice (250 issues, then 120 with fewer periods: no stale row). **Data checks: all 29 (DC01-DC29), done 2026-11-04.** The extraction now also reads `parent`, `resolutiondate`, `fixVersions`, `issuelinks` and the configured fields (story points, business value, PI, team, sprint, target dates): `Issues` has 22 columns, plus the raw table `IssueLinks` (Blocker, Blocked; outward "Blocks" links only) and `KeyChanges` (changelog items with `field="Key"` or `fieldId="issuekey"`). New Config rows: `FeatureTypes`, `SubTaskTypes`, `PointsRequiredTypes`, `ParentRequiredTypes`, the seven field ids, `BlocksLinkType`, `FillRateThreshold`; new sheet `PIs` (PI, Start, End). All client values stay empty in the repo. `flow\checks.py` is the Python reference (`run_checks`, `fill_rates`, `CheckSettings`); `flow\check_formulas.py` holds the Excel formulas (copies of raw fields, helper columns, 28 flags in `ItemFlow`, `IssueLinks` and `FillRates` columns, `DataChecks`). DC25 is counted on `FillRates` (field x team x issue type). **A check whose Jira field is not configured is `not computable` (Ready = 0) and flags nobody.** `mockjira\anomalies.py` holds hand-written clean cases and one anomaly per check; `flow_excel.feature` compares ItemFlow flags, `DataChecks`, `FillRates` and the known anomalies with the reference, and a fourth scenario checks the empty-field case. **Interpretations (all PROPOSALS to validate with AL and the teams; the document is silent or short on them):** DC06 also covers CURRENT_MISMATCH; DC28 compares consecutive intervals with `Transitions` and `IsFree`; DC01 applies to the Team workflow and `ParentRequiredTypes`; DC02 flags only when the parent is in the extract and its type is not in `FeatureTypes`; DC03 flags sub-tasks of a team that has more than one parent type; the current sprint is the active one, else a future one, else the last closed one; DC15 uses Jira status category `indeterminate` with no sprint; DC16 counts the sprints in the Sprint field (3 or more); PI format is `PI#[0-9]{2}-[0-9]{4}` and a malformed PI counts as empty for DC17 and DC18; DC18 uses the `PIs` window and the dates of the children's sprints; DC19 applies to features whose category is not Backlog or Removed; DC20 checks end < start and child dates outside the parent's dates; DC22 flags the blocked issue when the blocker's sprint ends after the blocked issue's sprint starts; DC23 flags the side without sprint on a cross-team link; DC25 uses `FillRateThreshold`; DC27 counts the key changes. **Not done:** (none of the data checks); **progress bar, cancel and sliced extraction (done 2026-11-04)**: the extraction is a job of steps in `modExtract` (`ExtractJobStart`, `ExtractJobStep`, `ExtractJobFraction`, `ExtractJobLabel`, `ExtractJobAbort`; `ExtractFlowDataTo` runs the same steps in a loop, so there is one code path); the Dashboard button runs one step per `Application.OnTime` slice (`modActions.BtnExtractAndCompute`, `ExtractSlice`), moves the bar (`pbBack` / `pbFill`, status cell named `dashStatus`, cell A3), and `BtnCancelExtraction` stops it before the raw sheets are written (the earlier results stay); `AbortExtraction` unschedules a pending slice. The search has no known size, so its share of the bar (30 %) is an estimate. The build of the flow sheets is one slice (no cancel during it). Known limits: while a slice runs (one HTTP call) Excel is busy; a pending `OnTime` reopens the workbook if it is closed during a run (`AbortExtraction` is for that). The mock has `delay=` (seconds per call) for the scenarios; the two scenarios on the button in `flow_excel.feature` check the bar, that Excel answers between slices, and the kept results. Also done: (done 2026-11-04: the daily Flow Load curve `flow_load_daily`, last chart, one ChartData row per day via `ChartSpec.Axis`; and the 80-100 % band on Predictability via `ChartSpec.Band`, a PROPOSAL; markers FLB33-34); the API token is typed in `Config` and stays in the workbook (clear it before sharing; README says so). Review findings fixed in a second pass: `BaseUrl` must be https (local fake Jira excepted), Windows credentials go to any server only when a proxy is set, the search asks only the 5 fields + WorkTypeField, the selector list moved beside the widest chart block. Review findings consciously NOT applied: recalculating twice after a build (speed only, not timed), duplicated sheet helpers (modExtract/modFlowBuild), missing `Cov` on a few small branches of modFlowBuild, `T_FlowBuild` cleanup on failure, ChartData block layout written both in `dashboard_spec.block_top` and in VBA (the scenario would catch a drift), some Python test hygiene and dead-code items |
| `out\` | Generated workbooks: `vba_tests_v11.xlsm` is the latest (2026-11-05, with `modJiraConfig`: 375 + 8 network OK, 0 KO, 189/189 markers) |
| `archive\` | Old code kept for history: `probe*.py`, `peek*.py`, `repro_*.py`, `bootstrap_test*`, `vba\patch1.py` ... `patch7.py` (already applied; do not re-run), old workbooks |

MD5 of the current modules (LF line endings). Unchanged since the original handoff: `modCoverage`,
`modFlow`, `modBackground`, `modButtons`. Changed (E2, E3 and the proxy work, 2026-11-03/04):
- `modParams` (new): reads the sheets; `ConfigValue`, `IsYes`, `ProjectKeys`, `ProgramProject`, `TeamOfProject`,
  `ParseMoment` / `AsOfFromConfig` (the UTC as-of date of `Config.AsOf`, parsed by hand: `CDate` depends on the language).
- `modExtract` (new, E4 stage 2): `ExtractFlowData` writes the RAW sheets `Issues` and `StatusIntervals` from Jira
  (search, bulk changelog, `modFlow.BuildStatusIntervals` once per workflow: the Program follows Feature, the others
  Team). VBA writes no figure: the formulas do.
- `modText`: `TrimWhite` (like PowerShell `.Trim()`).
- `modJira`: `JiraConfig` trims URL, email, token and normalizes the proxy (`NormalizeProxy`: `http://host:port/` or
  `host:port`); `JiraConfigFromSheet` / `JiraConfigFromWorkbook` read the `Config` sheet; `CheckStatus` adds a
  hint naming the Config cells to check for 407, 401 and "no answer"; `ProjectsJql` takes the optional Filter;
  `ExtractionJql` / `ExtractionJqlFromWorkbook` build the search from the `Projects` and `Config` sheets.
- `modHttp`: **retries every 5xx** (not only 503 with `Retry-After`) and 429, like the PowerShell script; timeouts
  60 s (connect, send, receive); sends `Accept-Charset: utf-8`; `ProxyWindowsAuth = No` now sets
  `AutoLogonPolicy_Never` (the Windows session is never sent to the proxy). `docs\jira-cloud-api-vba-reference.md`
  still says "503 only with Retry-After" and 120 s: it is not regenerated; this note supersedes it.
- `modFlowBuild` (new, E6): `BuildFlowSheetsIn` rebuilds ItemFlow, DailyLoad, Metrics, Controls, ChartData and the charts from the sheets `FlowFormulas` and `ChartSpec` (written at build time by `flow\blueprint.py`); sizes follow the data, nothing of an older run stays; `ClearFlowSheets` empties everything. `modActions` (new): `ExtractAndCompute` and the button macro `BtnExtractAndCompute`. The extraction empties the old results only once Jira has answered.
- Review fixes (2026-11-04): `AsOf` typed as a date in the cell is accepted (`AsOfFromConfig`, and the formula of `Controls`), an empty `AsOf` or a `FlowLoadMode` other than A/B gives `#N/A` instead of a guess, `ParseJiraDate` never raises, changelog paging and `Retry-After` are capped, `MaxRetries` < 0 is clamped, `JsonStringify` writes `0.5` not `.5`, `T_JsonErrors` really calls the parser.
- Code review of `src/` and `vba/` (2026-11-04, second pass): see section 8b for every finding and what was done.
- `modLog` (new, 2026-11-04): the local log of the runs, `%LOCALAPPDATA%\jira2dashboard\logs\jira2dashboard.log`
  (`%TEMP%` if not set). One line opened, written and closed at a time (survives an Excel crash): `RUN` at the start
  (Excel version and build, workbook name) and at the end, `BEGIN`/`END` around each step (search page, changelog
  chunk, raw sheets, flow sheets), `ERROR` with `vbObjectError+n`. Never the token, e-mail, URL, JQL or an answer of
  Jira: errors 602 and 607 (they quote the answer) log their number only; `tests\unit\test_vba_log_guard.py` refuses
  such arguments in any call to the logger. Rotation above 1 MB to `.1.log`. Never raises. Config `LogLevel`: Steps
  (default) or Errors (no BEGIN/END). A run that does not end leaves a flag in HKCU (`SaveSetting jira2dashboard`,
  section `Log`; the tests use `LogTests`): the next run writes `RUN previous run did not end (started ...)`.
  `LogRunEnd` writes `end ok|cancelled|stopped` and clears it. `T_Log` (markers LOG01-LOG13); the button scenario
  checks the lines of its run. The VBA test runner writes to
  `%TEMP%\jira2dashboard_vba_tests.log`, never to the real log.
- `modJiraConfig` (new, 2026-11-05): button **Read the Jira configuration** on the Dashboard
  (`modActions.BtnReadJiraConfiguration`, cell O1). Reads Jira read-only for the keys of `Projects` and writes proposal
  sheets, never the parameter sheets: `JiraIssueTypes` (Project, IssueType, Workflow, Subtask, HierarchyLevel,
  ProjectStyle), `JiraStatuses` (Project, IssueType, Status, JiraCategory), `JiraWorkflows` (the columns of `Workflows`
  + JiraCategory), `JiraTransitions` (the columns of `Transitions` + JiraTransition), `JiraFields` (Id, Name, Custom,
  SchemaType, SchemaDetail). Calls: section 8c (`project/search`, `project/{key}/statuses`, `workflows/preview`,
  `field`; `workflowscheme/project` and `POST /workflows` left out: they need Administer Jira, repeat the preview and
  miss team-managed projects). Proposals: to do -> Backlog, done -> Done, in progress -> empty; `IsFree = 1` for the
  target of a GLOBAL transition; `Order = 1` for the target of the INITIAL one; `IsStart`, `IsEnd` and `Type` empty (human
  decisions). A 401, 403 or 404 on the preview (no right to read workflows) leaves `JiraWorkflows` and
  `JiraTransitions` empty and says so in the status cell; the other sheets are written. Sheets are written only at the
  end: a cancel or an error keeps the proposals of the run before. A project whose issue types follow several
  workflows gives one `JiraIssueTypes` row per type with its own workflow (tested); workflows are told apart by id,
  so two project-scoped workflows with the same name both show. One go (about two calls per project), no slices;
  Cancel works during a wait for Jira. The log gets step names and counts only. `T_JiraConfig` (markers JCF01-JCF23),
  `M_JiraConfig` (round trip with the fake Jira), `tests\features\jira_config.feature` (two `@excel` scenarios; the
  first one also checks that the Dashboard button is at O1 and runs this macro). The fake
  Jira answers these calls (`mockjira\config.py`), built from the parameters, so the proposals must give them back.
- `modTests`: `T_FlowBuild` (builds twice, smaller the second time), parameter tests, `T_TrimWhite`, `T_ConfigSheet`, and the `M_` tests against the fake Jira.

| Module | MD5 |
|---|---|
| modCoverage.bas | 0b03cf2cb8d0660ed22beadabc9655a4 |
| modLog.bas | 8bba9ce1bb95bb061acb5b911fcb5752 |
| modJson.bas | e051a7775e419b2b49d2de6014a8719e |
| modText.bas | 2f23c0bc24c95cae731f331dbe42c047 |
| modHttp.bas | 849a7542964572acbece130c0d9bf086 |
| modJira.bas | 3e91019bff4c37532ad75d3beca54096 |
| modFlow.bas | 69f021411e3989de981f8bbde8ec5d73 |
| modParams.bas | 718b90c97703dde1073480eb4cee0d3c |
| modExtract.bas | d78be88ec4c062903dbfaf42586b9fa6 |
| modJiraConfig.bas | b4529efc36a4d0323c9508a1b7bccc56 |
| modFlowBuild.bas | bb2fbc8c73909f2ac82cf34b27c6ca08 |
| modActions.bas | 09b9bacac7238a8704fba4e8e51ba120 |
| modBackground.bas | 1d098ce38925b8dc55bcf54404cf314f |
| modTests.bas | 941f7fada66f76266b2cc55a9a76e94b |
| modButtons.bas | 3bbe95b810cc0722eabed5c2bdbabbb2 |

The Python files are no longer hashed (they were moved and edited; the code is covered by tests). If you
change a `.bas` file, update the MD5 table above and re-run `scripts\run.ps1 test`, `test-excel`,
`test-network` and `vba-tests`. `docs\vba-modules-tested.md` and `docs\cookbook_python.md` are not
regenerated any more.

## 5. Environment (dev machine)

Windows 11 x64, Python 3.14.7 (64-bit, `py -3.14`), pywin32 312, pytest 9.1.1, Excel 16.0 Microsoft 365
**French UI**. "Trust access to the VBA project object model" is **currently ON** (AL enabled it for the
tests; it can be switched off). An antivirus is active.
In `cmd.exe`: `chcp 65001` and `set PYTHONIOENCODING=utf-8` before running the scripts.

Client machine: Windows 11, PowerShell 5.1, Python 3.14, pywin32 312, licensed Excel M365, offline Copilot
GPT-5.6, Jira Cloud via a proxy (manual proxy address, Windows auth), an existing working PowerShell
extraction script (Copilot must read its proxy settings there; never copy the token).

## 6. Test status (2026-11-05, with the button Read the Jira configuration)

- `scripts\run.ps1 vba-tests --out vba_tests_v11.xlsm`: COMPILED_OK, **375 offline + 8 network OK, 0 KO**,
  coverage **189/189 markers = 100 %** (`modJiraConfig`: JCF01-JCF23; `modLog`: LOG01-LOG13; review markers: section 8b; earlier: E2: `modParams`; E3: `TrimWhite`; proxy work: `Config` sheet reader, proxy
  normalization, status hints; projects and filter; E4: `modExtract`, `ParseMoment`).
  The `M_` tests against the fake Jira are not in these counts: they run from `scripts\run.ps1 test-excel`.
- **Flaky Excel sessions while Office updates itself:** if `OfficeClickToRun` / `OfficeC2RClient` runs, Windows
  Restart Manager asks Excel to close and restart (clean exit, code 0; a new `EXCEL.EXE` without `/automation`
  appears and shows Document Recovery). A long hidden session can then vanish ("RPC server unavailable") and
  `test-excel` fails on the Excel scenarios. Wait for the update to finish and re-run. Not a code defect.
  **2026-11-04: the same symptom was a real Excel crash** (`0xC0000005`, event log "Application Error"), seen before
  and after an Office update; Office's crash recovery can end the process with code 0 and start a new `EXCEL.EXE`, so
  the diagnosis above may have been the same crash. Check the exit code first (section 8b).
- `scripts\run.ps1 check-machine`: **21 checks, 0 failed**.
- `scripts\run.ps1 coverage`: **623 passed**, 98.9 %. `test-network`: 1 passed. `build`: COMPILED_OK.
- `scripts\run.ps1 test-excel`: 34 of 35 on the last run (2026-11-05, with the two scenarios of `jira_config.feature`
  and `M_JiraConfig` in the fake Jira scenario); the button scenario lost its Excel right after `END search page 30:
  150 issues` (**open defect**, section 8b; run alone it also crashes: 7 out of 10 that night). Earlier: 9 module runs out of 10 green with a fresh Excel
  per scenario.
- Jira itself was never called (no access to the client Jira; the Jira plugin has no site configured).
  Jira behaviour is tested with fake responses (`FakeAdd`). Proxy: to verify at the client.

## 7. Findings from testing (all are in the docs; do not regress)

| Trap | Rule |
|---|---|
| pywin32 named COM args silently ignored (`Worksheets.Add(After=x)` put the sheet first) | Positional args only; `None` to skip one |
| Naive `datetime` shifted by the local time zone; `date` rejected | `to_excel_serial()` |
| `NumberFormat` set from Python is read in French (`0.00` -> 0.06, `yyyy-mm-dd` -> `yyyy-11-dd`) | `put_en_us(obj, "NumberFormat", fmt)` (LCID 1033). VBA `.NumberFormat` is fine |
| `.Text` depends on format and column width | Read `.Value2` |
| Excel converts written text (`true`, `1.10`, `2026-10-01`, `1/2`, `007`, `50%`, `#N/A`, `=1+1`) | Leading apostrophe on every text: Python `as_cell()`, VBA `SafeCell()` |
| New sheet `CodeName` empty until `VBProject` is touched | `sheet_component()` |
| Unsaved one-sheet workbook is named `Feuil1` | Save before `Application.Run` |
| `app.International(3)` returns a tuple | `app.International[2]` |
| Excel needs about 4 s to exit after the last COM reference | `excel_session()` 5 s grace, then ends its own PID only |
| `Ping` passes even if another module does not compile | `compile_vba_project()` (VBE control 578) in a child with timeout |
| VBA: `True & ""`, `Str$(True)` give "Vrai" | `IIf(b, "True", "False")` |
| VBA: `On Error Resume Next` leaves a stale `Err` | `Err.Clear` before leaving |
| `\u00e9` escapes became real accents when written by tools | Keep VBA ASCII; never retype modules; extract with MD5 |
| PowerShell started from VBA blocked by the antivirus | Option B: user runs the script; Excel only watches the export folder (`modBackground`) |
| Functions / Subs with parameters not in Alt+F8 and not usable on buttons | `Public Sub Btn...()` without args in `modButtons` |
| 14 s runs without feedback: user clicks again and clears results | Busy flag, row-by-row results, progress bar (`pbBack`/`pbFill`, label J10), status bar |
| JSON parser speed | ~170 KB/s (635 KB in 3.7 s) |

## 8. Open decisions (AL must validate)

1. Status categories: **decided** Nouveau and Depriorisé = Backlog, Terminé = Done, Annulé = Removed;
   transitions, start (first En analyse, reset after Backlog) and end (Terminé) in
   `docs\safe-flow-metrics-and-data-checks.md` section 1. Active: En analyse, En developpement, En test;
   Wait: Pret, Planifie, A deployer (all decided). `FieldMapping` validated 2026-11-03. **Still open:** nothing on the status and field decisions.
   `Livre` = released (no use for now); PI format `PI#XX-YYYY` checked by DC29; custom field 11339 not needed.
2. Rovo: **decided** one agent "ART Data Quality" with subagents (`docs\rovo-guide.md` section 4). Superseded
   on 2026-11-05: one agent "ART Flow Steering" with three jobs and no subagents (`rovo\rovo_train_steering.md`).
3. Workflow validators: written as a proposal to validate with the teams; Depriorisé sets no resolution (decided).
4. Sandbox of 20 seeded issues: table written (`docs\rovo-guide.md` section 7), to build in Jira.
5. Placement `.github/copilot-instructions.md` + `docs/` (check the client Copilot version).
6. **Validated:** `RovoCompare` and `RawFields` sheets, Confluence page "ART data rules" as single source of rules.
7. Not tested on an **English** Excel yet (run `check_reference.py` once there).

## 8b. Code review of `src/` and `vba/` (2026-11-04, `/engineering:code-review`, two reviewers, every finding verified)

All findings were treated. **Fixed** means a test or a scenario covers it, or the change is in the code path the scenarios run.

**Python (20 findings)**

| # | Finding | Disposition |
|---|---|---|
| 1 | Excel orphaned when a property set fails between `DispatchEx` and the `try` | Fixed: pid read at once, properties set with `call_with_retry`, Excel quit and killed on failure; `_quit_quietly` is bounded |
| 2 | `vba-tests` could exit 0 with no result or missed markers | Fixed: exit 1 when no test ran, no marker, or (not offline) a marker is missed; scalar `UsedRange` handled |
| 3 | A check that reads a type list (`FeatureTypes`...) is never "not computable" when the list is empty | Not applied: the defaults are filled and the lists are documented; an empty list is a choice ("no such type"). Proposal for AL: make an empty list "not computable" like an empty field id. `blueprint.check_rows` now asserts at most 3 needs |
| 4, 11 | DC18 and the daily load disagree with a time of day in `PIs` / `Periods` | Fixed: `Periods` and `PIs` must hold whole days (validated in `params/model.py`) |
| 5 | The oracle has no row for a team without issue (the chart falls back on the workbook's own figure) | Open, test gap only: no wrong figure is known. To do with the next change of the metrics reference |
| 6 | Type lists with two spaces or a non-breaking space differ between Python and Excel | Fixed: `TEXTJOIN/TRIM/TEXTSPLIT` in `check_formulas._list` |
| 7 | Work type vocabulary (Feature, Enabler, Defect, Autre) is in the code | Not applied: the four buckets are the SAFe Flow Distribution of `docs\`; asking AL whether a client may need other labels |
| 8 | `compile_vba_project` reported OK when the control was disabled | Fixed: raises |
| 9 | The build never checked VBA access | Fixed: `assert_vba_access` right after `new_workbook` |
| 10 | A period starting after `AsOf` shows Velocity 0 | Not applied: decision for AL (the document says "not computable" only for a missing figure). Proposal: NC for periods that have not started |
| 12 | DC19 uses the category of the status, Excel the category of the open interval; DC18 crashed on a sprint with a start and no end | DC18 fixed. DC19: documented, the two differ only when DC06 flags the issue (CURRENT_MISMATCH) |
| 13 | `CurrentCategory` and `SubParentTypes` scan the whole table for every row | Not applied: to measure on a client-size extract first. Proposal: a helper column on `StatusIntervals` |
| 14 | `save_as_xlsm` deletes the target before `SaveAs`; `--out` can leave `out\` | `--out` fixed (bare file name ending in `.xlsm`). Delete-before-save not applied (the callers close the workbook after `SaveAs`) |
| 15 | Mock: bad `Content-Length`; control endpoints before the credentials check | Content-Length fixed (400). Control endpoints: not applied, the mock is a local test tool |
| 16 | The mock banner prints its token | Not applied: it is the fake account to type in the workbook, never a real token |
| 17 | The "no token in params" guard only looks at the `ApiToken` row | Fixed for `BaseUrl` (no `user:token@host`) |
| 18 | Dead or latent helpers (`vba_string_literal`, `end_statuses`, test mirrors) | Not applied: kept, harmless |
| 19 | The test `retry()` hides real errors as "Excel stayed busy" | Fixed: the last error is chained |
| 20 | `known_anomalies` fails with an unclear error when fewer than 3 projects | Fixed: clear message |

**VBA (25 findings)**

| # | Finding | Disposition |
|---|---|---|
| 1, 12 | The Periods sheet or an old Excel was only found after the extraction, with the old results already cleared | Fixed: `ValidateBuildInputs` (sheets, plan, teams, periods, Excel 365 probe) runs before the first call to Jira |
| 2 | A slice still scheduled reopens the workbook after it was closed | Fixed: the build adds `Workbook_BeforeClose` (calls `AbortExtraction`); the delivery fixture checks it exists |
| 3 | `OnTime` with a bare procedure name | Fixed: qualified with the workbook name |
| 4 | A network error (status 0) was not retried | Fixed: two retries, then the error (`T_HttpNetworkAndCancel`) |
| 5 | Cancel not read during a wait for Jira | Fixed: `gHttpCancel` ends the waits; `HttpRequest` raises 610; `ExtractSlice` shows "Cancelled" |
| 6 | `IsLocalUrl` accepted `localhost.evil.example` and `?x=.invalid` | Fixed: the host is parsed (4 tests) |
| 7 | The job kept every page and the Authorization header after the run | Fixed: `ExtractJobAbort` releases everything and is called at the end |
| 8 | `IssueTypes` cell by cell; no error without a Type column | Fixed: one read for 2 rows or more (the one-row case needs the single-cell read: the reviewer's claim was wrong there); a missing column raises |
| 9 | Daily Flow Load point scanned the whole table | Fixed: `COUNTIFS` and `SUMIFS` |
| 10, 11 | The build left another sheet active, forced `ScreenUpdating`, and the raw sheets were written with automatic calculation | Fixed |
| 13 | `ExtractAndCompute` had no busy guard and no cleanup | Fixed: error 611 while a run is active; the status bar is reset |
| 14 | `gHttpLog` grew in production | Fixed: filled only once a test has started it (`FakeReset`); production never starts it |
| 15 | Full jitter could give an instant retry | Fixed: half jitter |
| 16 | A proxy login page (200, HTML) gave a bare parser error | Fixed: `RequireJson` explains and shows the start of the answer |
| 17 | Redirects followed with the Authorization header | Not applied: could not be checked against WinHTTP here and disabling redirects could break a client proxy. To check at the client |
| 18 | The API token is in clear text in `Config` | Not applied: known design (README: clear it before sharing). Proposal: ask for it when the cell is empty |
| 19 | A bad `\u` escape gave error 13 | Fixed |
| 20 | No guard for a repeated page token in the search | Fixed |
| 21 | Duplicated "first day" loops, two sheet helpers, the 2000-page limit twice | Not applied: low value, no defect |
| 22 | Test-only helpers in delivered modules (`IssueFieldIds`, `IsAllowed`...) and `modBackground` | Not applied: used by the tests; `modBackground` is not delivered |
| 23 | `NumberOf` / `TextOf` read numbers and Booleans with the locale of the PC | Fixed |
| 24 | Cancel pressed during the build was ignored silently | Fixed: says it is too late |
| 25 | `AsOf` typed as 5 became 1900-01-05 | Fixed (error 604 under 2000-01-01) |

**New in the code after the review (for the next reader).**
- `modFlowBuild.ValidateBuildInputs`: `modExtract.ExtractParameters` calls it before any call to Jira. It checks the sheets, the plan, the teams, the periods and Excel 365 (`RequireDynamicArrays`, error 708). `IssueTypes` raises 709 without a Type column.
- `modHttp.gHttpCancel`: the Cancel button sets it, the waits end and `HttpRequest` raises 610. A network error (status 0) is retried twice (`MAX_NET_RETRIES`).
- `modJira.RequireJson` (error 607) on the search and the changelogs; error 606 when Jira sends the same page token twice.
- `modActions`: error 611 when `ExtractAndCompute` runs during a sliced run. `OnTime` uses `'<workbook>'!ExtractSlice`.
- `build_workbook.py` writes `ThisWorkbook.Workbook_BeforeClose`, which calls `modActions.AbortExtraction`.
- Coverage markers: 153 (new: FLB35, HTP08, HTP09, JIR24, JIR25, JSN19, JSN20, PRM25).

**OPEN DEFECT: Excel crashes during the extraction, much more often after several extractions in the same Excel
process (2026-11-04).** With one Excel for
the whole module, about 2 runs out of 5 of `flow_excel.feature` lost their Excel ("RPC server unavailable"). It is a
real crash: `EXCEL.EXE`, exception `0xC0000005`, faulting module "unknown" (Application event log, event 1000).
What is known:
- **Where (11 crash dumps, `%LOCALAPPDATA%\CrashDumps\EXCEL.EXE.*.dmp`, all identical):** a DEP fault on Excel's main
  thread. VBA's engine (`VBE7.DLL`, after `call rdx`, called through `oleaut32!DispCallFunc`) calls a small piece of
  valid x64 code that it generated at run time, in heap memory that is not executable at that moment. One VBA level on
  the stack; no WinHTTP, ADODB, Scripting or nested `DoEvents` frame. Panda Security (WatchGuard) hook DLLs
  (`psninjhook*`) are loaded in every dump.
- **When:** with **a fresh Excel per scenario, 1 module run out of 10 failed** (the button scenario lost its Excel
  during the search; exit code 0, no event 1000 and no dump that time). With one Excel for the module: about 2 out of 5, mostly in the button scenarios (the 5th and 6th extractions of the
  process), in the changelog slice (about 10 s of pure VBA: parse a 216 KB answer, then 740 rows with late-bound
  `Scripting.Dictionary` calls). More crashes when the machine is loaded.
- Seen on Excel 16.0.20527.20112 and 16.0.20527.20150; the code before the review crashes as often (2 out of 4).
- **Located by the local log (2 crashes, 2026-11-04, button scenario, fresh Excel):** the last line of the run is
  `END search page 30: 150 issues`; `BEGIN changelogs from issue 1 of 150` never comes. Excel dies between the end of
  the last search page and the start of the next slice: the progress update (`DoEvents`), the `OnTime` scheduling or
  Excel's call of the next slice. With the dumps (9 of 11 entered VBA from Excel's main loop), the likeliest moment is
  Excel starting the `OnTime` macro, not the work inside a slice. Why only at that transition is not known.
- **Bench of the button scenario alone (fresh Excel, 2026-11-04 evening):** the committed code crashed in 3 runs out
  of 4; a variant that runs the whole job inside one `OnTime` call (no `OnTime` between the steps) crashed in 1 out
  of 4. **Every crash has the same last line, `END search page 30: 150 issues`**, so the `OnTime` entry is not the
  cause. Two temporary log lines then placed 3 more crashes: once inside the loop that lists the 150 issue keys (the
  first statements after the last page), twice inside `ShowProgress` right after it. Excel dies 0.2 to 0.5 s after
  the last search page, at different ordinary VBA statements: a moment, not a statement of our code. This fits the
  dumps (VBA's engine runs code from memory that is not executable) and points to the state of the process memory:
  the security software's anti-exploit (Panda hooks in every dump) or a VBA engine bug. No code change tried so far
  removes it.
- Same night, the run flag of `modLog` worked: after a crash of the button scenario, the next run wrote `RUN previous
  run did not end (started ...)`. Once, the Excel of `test-network` (the VBA suite with network) died too, with
  `0xC0000028` (bad stack during an exception unwind) in `ntdll.dll`: probably the same defect seen from another side
  (an exception crossing code generated at run time); the rerun passed.
- Tried without effect: slower COM polling (kept), no manual-calculation wrapper, POST body kept alive and
  `req.Abort` on cancel (kept as hardening), synchronous WinHTTP, keeping the last WinHTTP request alive.
- Office's crash recovery can end the process with code 0 and start a new `EXCEL.EXE` in the same millisecond.

Causes (ranked, PROPOSALS): (1) the anti-exploit module of the security software removes the execute right from VBA's
code memory (documented by other vendors for Excel macros, not for Panda); (2) a VBA engine bug when the process has
grown; (3) COM calls of the test during `DoEvents` (unlikely: never on the crashing stack).

What was done: the Excel scenarios now open a fresh Excel per scenario (`flow_workbook` fixture; `J2D_ONE_EXCEL=1`
gives back one Excel for the whole module, the stress setup to test a fix: with a fresh Excel per scenario, 5 green
runs prove nothing), and the scenario "Un second calcul..." makes its own first calculation (250 issues) before the
second one (120 issues). The local log (`modLog`) now writes BEGIN/END around each step: after a crash, the last BEGIN
without its END names the step. Lead dropped: AMSI runtime scanning of VBA (Office scans low-trust files, a Trusted
Location is exempt): from a Trusted Location the changelog loop was as slow (837 rows in 10 s), so no sign of it.

**Delivery risk:** a user who runs several extractions in the same Excel session may hit it; the client's security
software may decide. Next steps (PROPOSALS, for AL): run the button several times in one Excel session at the client
before delivery; ask IT for a Panda anti-exploit exclusion (or audit mode) for `EXCEL.EXE` on the dev PC and run the
old one-Excel module 5 times; with admin rights, full-memory crash dumps (`LocalDumps\EXCEL.EXE`, DumpType 2) to read
the page protection; fewer late-bound calls in the changelog loop (8 to 13 s for about 1,000 rows); at the client,
read the last BEGIN without END of the local log after a crash (study: `C:\dev\reports\Journalisation des erreurs
VBA.md`, outside the repo).

## 8c. Jira configuration API (read-only, to fill the parameters)

Found 2026-11-05 (Context7, Jira Cloud platform OpenAPI spec). In Jira a project's **workflow scheme** maps each issue
type to a workflow: in one Team project, Bugs may follow another workflow than Stories. The workbook today picks the
workflow by project role only (Program = Feature, Team = Team, `docs\safe-flow-metrics-and-data-checks.md` section 1).
Scenarios for a workflow per issue type are written, not bound to steps yet: `tests\features\workflow_by_type.feature`.

All calls below are read-only, made at the client from the client machine; their answers stay in the workbook.
**Done 2026-11-05:** the button Read the Jira configuration (`modJiraConfig`, section 4) makes the calls of rows 1, 2,
5 and 6 and writes proposal sheets. The shapes are documented, not seen on a real Jira: verify at the client.

| Need (parameter) | Call | Permission | What to read |
|---|---|---|---|
| `Projects`: id, issue types, project style | `GET /rest/api/3/project/search?keys=AAA&keys=EEE&expand=issueTypes` (offset paging, `isLast`, max 100, up to 50 keys) | Browse projects | `values[].{id, key, name, style (classic/next-gen), issueTypes[].{id, name, subtask, hierarchyLevel}}` |
| Statuses **per issue type** | `GET /rest/api/3/project/{key}/statuses` | Browse projects | array of `{id, name (issue type), subtask, statuses[].{id, name, statusCategory.key (new/indeterminate/done)}}` |
| Workflow **of each issue type** | `GET /rest/api/3/workflowscheme/project?projectId=10000&projectId=10001` | Administer Jira | `values[].{projectIds, workflowScheme.{id, name, defaultWorkflow, issueTypeMappings {"<issue type id>": "<workflow name>"}}}`; team-managed projects are left out; the default scheme has no id |
| Statuses **and transitions** of workflows | `POST /rest/api/3/workflows`, body `{"projectAndIssueTypes":[{"projectId":"10000","issueTypeId":"10004"}]}` or `{"workflowNames":["..."]}` | Administer Jira, or Administer projects / View (read-only) workflow | `statuses[].{id, name, statusReference, statusCategory (TODO/IN_PROGRESS/DONE)}`, `workflows[].{name, statuses[].statusReference, transitions[].{type (INITIAL/GLOBAL/DIRECTED), toStatusReference, links[].fromStatusReference}}` |
| Same, read-only, per project | `POST /rest/api/3/workflows/preview`, body `{"projectId":"10000","issueTypeIds":["10004","10001"]}` (max 25 ids) | Administer projects or View (read-only) workflow | as above, plus `workflows[].queryContext[].{project, issueTypes}`: the issue types of the project that use each workflow |
| Field ids (Config) | `GET /rest/api/3/field`, `GET /rest/api/3/field/{id}/context` | Browse | see `docs\jira-cloud-api-vba-reference.md` section 1 |

How the answers fill the parameters: `queryContext` (or `issueTypeMappings`; unmapped types follow
`defaultWorkflow`) gives issue type -> workflow; the statuses of a workflow give the `Status` column of `Workflows`;
`statusCategory` gives a first proposal of `Category` (TODO -> Backlog, DONE -> Done, IN_PROGRESS -> Active or Wait,
to decide); `Order`, `IsStart`, `IsEnd` stay a human decision. Each link of a `DIRECTED` transition gives a
`Transitions` row From -> To (`Type` forward/back/exit to decide); a `GLOBAL` transition (from any status) means
`IsFree = 1` on its target; `INITIAL` gives the creation status. Without admin rights, `project/{key}/statuses`
already shows whether two issue types of a project use different statuses.

Traps: `GET /rest/api/3/workflow/search` was removed on 2026-06-01 (use `GET /rest/api/3/workflows/search` or the
calls above); team-managed projects are not in `workflowscheme/project` (use `workflows/preview`).

Rovo side (2026-11-05): `rovo\rovo_settings_analyser.md` configures a Rovo Studio agent "Settings Analyser" that
lists, for one project, the work item types, the total per type, the statuses per type and the custom fields, with
the headers of the `Jira*` sheets, to compare Jira (as Rovo sees it) with the workbook. Rovo agents read Jira with
JQL search only (other tools: verify in Studio): workflows and transitions are not visible to them, counts carry
their JQL. Not tried on a real site yet.
Also in `rovo\` (2026-11-05, from the research report `C:\dev\reports\Agents Rovo Studio et métriques.md`, outside
the repo): `rovo_flow_metric.md` (flow metrics per work item type, each number labelled exact, approximate or sample;
no official figure comes from Rovo) and `rovo-guide.md`, the updated copy of `docs\rovo-guide.md` (which stays
frozen): Studio options, counts and limits, the DC checks as JQL (10 exact, 9 partial, 1 to test, 9 workbook
only), and the agent Flow Data Steward.
Steering with Rovo (2026-11-05, AL's choice: hybrid). Rovo cannot reproduce the dashboard on its own: extraction,
status intervals, metrics and charts stay in the workbook. The RTE publishes `Metrics`, `DataChecks` and the chart
images on a Confluence page.
- `rovo\confluence_flow_metrics_page.md`: the template of that page.
- `rovo\rovo_train_steering.md`: the agent ART Flow Steering, which replaces Flow Data Steward. It has three jobs
  (readiness, explain, steering), one `rovo\skills\<name>\SKILL.md` each. It reads the page and never computes a
  metric.
- `rovo\jira_dashboard_data_quality.md`: a Jira dashboard of saved filters `DQ DCxx`, built by hand.

The page template follows the column names of `METRICS_COLUMNS` and `DATACHECKS_COLUMNS` and the chart titles: update
it when they change. Still to check at the client: whether a custom skill can be attached to a Studio agent, and the
page-read test (`rovo_train_steering.md` section 6).

## 9. Suggested next steps

1. Ask AL which of section 8 to validate first.
2. Prepare the transfer to the client: list of the 7 md files, OneDrive or zip if mail changes them
   (MD5 mismatch), installation order (bootstrap, `check_reference.py`, `run_vba_tests.py --offline`).
3. At the client, with Copilot: step 1 Config + extraction (option A or B), then intervals, checks, metrics;
   each step with tests and coverage; AL pastes outputs back.

## 10. Useful commands here

```text
cd C:\dev\jira2dashboard
powershell -ExecutionPolicy Bypass -File scripts\setup.ps1 -Dev
.\scripts\run.ps1 test
.\scripts\run.ps1 test-excel
.\scripts\run.ps1 test-network
.\scripts\run.ps1 check-machine
.\scripts\run.ps1 vba-tests --out vba_tests_v6.xlsm
start "" out\vba_tests_v6.xlsm
```

If a workbook is open in AL's Excel, write to a new `--out` name (the open file is locked).
Never kill an `EXCEL.EXE` that the scripts did not start: AL often has his own Excel open.
