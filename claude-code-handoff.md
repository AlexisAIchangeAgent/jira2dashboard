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
| `src\jira2dashboard\mockjira\` | **Fake Jira Cloud** (E3), stdlib `http.server` on 127.0.0.1: `GET/POST /rest/api/3/search/jql` (token paging, `expand=names`), `POST /search/approximate-count`, `POST /changelog/bulkfetch`, `GET /field`, `GET /myself`; `410` on the removed `/search`; faults (429, 500/503 with or without `Retry-After`) injected from Python (`fail_next`) or over HTTP (`POST /__mock__/fail`); Basic auth `base64("email:token")` checked when credentials are set (plain-text 401 like Jira); usable as an HTTP proxy: absolute-form requests and a plain `CONNECT` tunnel, with an optional proxy challenge (`proxy_challenge="Negotiate"`: 407 until the request carries `Proxy-Authorization`; any value is accepted, so a real NTLM/Kerberos token is not checked). Modelled on the existing PowerShell extraction script (project keys, BV/SP/PI custom field ids, Basic auth, `/rest/api/3` paths). JQL understood: `project in (...)` plus restrictions joined by `AND`: `created` / `updated` (>=, >, <=, < with `"2026-10-01"` or `-30d` / `-2w`, counted from the dataset's as-of date), `issuetype`, `status` (= or `in`), parentheses, `ORDER BY key|created|updated`; anything else is a 400 that says so. Not simulated: TLS 1.2 inside the tunnel and the real NTLM/Kerberos exchange (verify at the client). Data = issue status histories (`dataset.py`; `issue()` builds hand-written cases, `generator.py` builds a seeded volume), so search and changelog never disagree. Run it for Excel: `scripts\run.ps1 mock-jira` (`http://127.0.0.1:8099` by default), then put that URL in `BaseUrl` |
| `tests\` | `unit\` (no Excel), `features\` + `steps\` (Gherkin, pytest-bdd), `integration\` (Excel) |
| `scripts\` | `setup.ps1`, `activate.ps1`, `run.ps1` (PowerShell 5.1) |
| `src\jira2dashboard\build_workbook.py`, `vba\modFlowBuild.bas`, `modActions.bas` | **Delivery (E6).** `scripts\run.ps1 build` makes `out\jira2dashboard.xlsm` (10 modules, no test module). Button **Extract from Jira and compute** on the Dashboard runs `modActions.BtnExtractAndCompute`: extraction, then `BuildFlowSheetsIn`. Scenarios of `flow_excel.feature` run this macro on the delivered workbook against the fake Jira and compare with the Python reference, twice (250 issues, then 120 with fewer periods: no stale row). **Data checks (done after E6, 2026-11-04): DC01, DC04, DC05, DC06, DC07, DC08, DC09, DC10, DC28** (the ones the extraction supports; the extraction now also reads `parent`, `resolutiondate` and the Jira status category key into `Issues`: columns `Parent`, `StatusCategory`, `ResolvedUtc`; DC01 uses the new `Config.ParentRequiredTypes`, default `Story,Task`, team projects only). `flow\checks.py` is the pure Python reference (`issue_checks`; DC06 also covers CURRENT_MISMATCH, DC28 compares consecutive intervals with `Transitions` and `IsFree`: both are proposals, the document only says "Check contains it"). Excel: `StatusIntervals` gets `PrevStatus` and `BadMove`, `ItemFlow` gets six 0/1 columns `DC06`.., sheet `DataChecks` (Id, Title, Severity, Items, Result; ids come from the sheet `ChecksSpec`, written by `blueprint.check_rows()`; `modFlowBuild.WriteCheckKeys`, marker FLB30). `mockjira\anomalies.py` adds nine hand-written cases (eight anomalies and one clean control, keys `<team project>-9001..9009`; `Issue.resolution_override` lets a case lack or have a resolution date); third scenario of `flow_excel.feature` compares ItemFlow flags and DataChecks with the reference and checks that each case fails exactly its checks (a mutation of the DC10 formula was verified to fail it). **Not done:** the other 20 data checks DC02, DC03, DC11-DC27, DC29 (they need extra Jira fields not in the extraction: sprint, team, PI, dates, links, or a list of feature types); progress bar / cancel during the extraction (it is synchronous, the status bar shows the steps); Flow Load daily curve and the 80-100 % band of Predictability (plan E5 text); the API token is typed in `Config` and stays in the workbook (clear it before sharing; README says so). Review findings fixed in a second pass: `BaseUrl` must be https (local fake Jira excepted), Windows credentials go to any server only when a proxy is set, the search asks only the 5 fields + WorkTypeField, the selector list moved beside the widest chart block. Review findings consciously NOT applied: recalculating twice after a build (speed only, not timed), duplicated sheet helpers (modExtract/modFlowBuild), missing `Cov` on a few small branches of modFlowBuild, `T_FlowBuild` cleanup on failure, ChartData block layout written both in `dashboard_spec.block_top` and in VBA (the scenario would catch a drift), some Python test hygiene and dead-code items |
| `out\` | Generated workbooks: `vba_tests_v6.xlsm` is the latest (293 OK, 132/132 markers; since the data checks: 290 + 8 OK, 133/133) |
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
- `modTests`: `T_FlowBuild` (builds twice, smaller the second time), parameter tests, `T_TrimWhite`, `T_ConfigSheet`, and the `M_` tests against the fake Jira.

| Module | MD5 |
|---|---|
| modCoverage.bas | 0b03cf2cb8d0660ed22beadabc9655a4 |
| modJson.bas | 281dc819db5da7f08a9a280c218dbf86 |
| modText.bas | 2f23c0bc24c95cae731f331dbe42c047 |
| modHttp.bas | 939654dc73f1b4aeb2248530efae7c2c |
| modJira.bas | f9d00a473758bb1f503d34ee35cc55c7 |
| modFlow.bas | d709521a710f21e3319df21038c8e463 |
| modParams.bas | c2cc338cf79d8ade69fc66add8d4dc96 |
| modExtract.bas | cee57d93f2acd808ce033d89c585321a |
| modFlowBuild.bas | d22dd25dae5dd617768605e318435350 |
| modActions.bas | 33fb4e5a39ed56bf5a3c6dd496c74456 |
| modBackground.bas | 1d098ce38925b8dc55bcf54404cf314f |
| modTests.bas | babec17d115a0942151adf23323e5944 |
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

## 6. Test status (2026-11-03)

- `scripts\run.ps1 vba-tests --out vba_tests_v6.xlsm`: COMPILED_OK, **290 offline + 8 network = 298 OK, 0 KO**,
  coverage **133/133 markers = 100 %** (E2: `modParams`; E3: `TrimWhite`; proxy work: `Config` sheet reader, proxy
  normalization, status hints; projects and filter; E4: `modExtract`, `ParseMoment`).
  The `M_` tests against the fake Jira are not in these counts: they run from `scripts\run.ps1 test-excel`.
- **Flaky Excel sessions while Office updates itself:** if `OfficeClickToRun` / `OfficeC2RClient` runs, Windows
  Restart Manager asks Excel to close and restart (clean exit, code 0; a new `EXCEL.EXE` without `/automation`
  appears and shows Document Recovery). A long hidden session can then vanish ("RPC server unavailable") and
  `test-excel` fails on the Excel scenarios. Wait for the update to finish and re-run. Not a code defect.
- `scripts\run.ps1 check-machine`: **21 checks, 0 failed**.
- `scripts\run.ps1 test`: unit tests green (no Excel); `test-excel` and `test-network`: Excel scenarios green.
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
2. Rovo: **decided** one agent "ART Data Quality" with subagents (`docs\rovo-guide.md` section 4).
3. Workflow validators: written as a proposal to validate with the teams; Depriorisé sets no resolution (decided).
4. Sandbox of 20 seeded issues: table written (`docs\rovo-guide.md` section 7), to build in Jira.
5. Placement `.github/copilot-instructions.md` + `docs/` (check the client Copilot version).
6. **Validated:** `RovoCompare` and `RawFields` sheets, Confluence page "ART data rules" as single source of rules.
7. Not tested on an **English** Excel yet (run `check_reference.py` once there).

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
