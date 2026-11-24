# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

**jira2dashboard** builds an Excel workbook (`.xlsm`) that extracts Jira data, computes the six SAFe flow metrics
(Velocity, Flow Time, Flow Load, Flow Efficiency, Flow Distribution, Flow Predictability) and shows them on a dashboard
with native charts. Python (pywin32 + stdlib at runtime) **builds** the workbook and its VBA; the delivered workbook then
works **without Python**: VBA reads Jira, builds the tables, formulas and charts.

Platform: Windows 11, PowerShell 5.1, Python 3.14, Excel Microsoft 365 (French UI on the dev machine). Runtime = pywin32 +
stdlib only; pytest, pytest-bdd and pytest-cov are in the same `requirements.txt` (one install, no `-Dev`; `pyproject.toml` has no `dev` group, and `pywin32`'s version is written in both files: change both). The repo is self-contained: child Excel processes import
`jira2dashboard` through `PYTHONPATH=<repo>\src` (no `pip install`).

Docs and code are in English. Conversation with the user (AL) is in French, short simple sentences.

## Commands

Always go through `scripts\run.ps1` (PowerShell 5.1, no venv activation needed); extra arguments reach pytest/Python.

```powershell
powershell -ExecutionPolicy Bypass -File scripts\setup.ps1        # create .venv (py -3.14) and install requirements.txt (pywin32 + test tools)
.\scripts\activate.ps1            # optional: activate .venv in this console (run.ps1 needs none)
.\scripts\bootstrap.ps1         # all in one: setup, check-machine, test, build the .xlsm, the .xlsx (scenario); -Full adds vba-tests + test-excel; log bootstrap.log
                                # not the full pass of a release: coverage, test-network and test-scripts stay to run apart
.\scripts\run.ps1 test            # fast unit + Gherkin tests, no Excel (the inner loop)
.\scripts\run.ps1 coverage        # same + coverage of the pure code; fails under 93 %
.\scripts\run.ps1 test-excel      # scenarios that drive Excel (marker `excel`, no network)
.\scripts\run.ps1 test-network    # Excel + network scenarios (marker `network`)
.\scripts\run.ps1 vba-tests       # build + compile + run the VBA test suite; --offline, --bisect, --out name.xlsm
.\scripts\run.ps1 check-machine   # 24 checks of the machine (Excel, VBA trust, what Excel writes stays, macro of the workbook)
.\scripts\run.ps1 diagnose        # paste-ready report when the macros are blocked (hashes, Mark of the Web, Office, security products)
.\scripts\run.ps1 mock-jira       # fake Jira on http://127.0.0.1:8099 (--anomalies: with the known data anomalies)
.\scripts\run.ps1 mock-workbook   # the filled copy out\manual_test\jira2dashboard4manualtest.xlsm and its Rovo CSV (--port, --workbook)
.\scripts\run.ps1 build           # delivery workbook out\release\jira2dashboard.xlsm (--out name.xlsm: to out\inter)
.\scripts\run.ps1 clean           # empty out\inter except the test caches and the logs (build and vba-tests do it first)
.\scripts\run.ps1 extract         # Python path without VBA: Jira -> jira2dashboard.xlsx (values + SVG charts)
.\scripts\manual-test.ps1         # manual try: fake Jira + pre-filled copy (mock-workbook), opened in a new Excel
.\scripts\stop-mock.ps1           # kill the fake Jira that start-mock.ps1 / manual-test.ps1 started
.\scripts\scenario.ps1            # unattended scenario: fake Jira, filled copy, Python extraction, checks, clean-up (-Excel: + test-excel)
.\scripts\run.ps1 test-scripts    # Pester 3.4 tests of scripts\*.ps1 (tests\ps), in process, with fakes; fails under 93 % command coverage
.\scripts\release.ps1 -Version v1.0.0 [-Publish]   # built workbook in out\release (no zip of the code); -Publish = GitHub release
```

Every `scripts\*.ps1` has a double-click launcher `scripts\<name>.cmd` (shared helper `_launch.cmd`: `-ExecutionPolicy Bypass`
for that run only, waits for a key only on a double-click or with `J2D_PAUSE=1`, `--help`, exit code kept). A new `.ps1` needs
one: copy `stop-mock.cmd` and change its `set` lines. `.cmd` files are ASCII with CRLF and their description line has no
`<>|&^%!` (cmd would take them for operators); `tests/unit/test_cmd_launchers.py` pins all this. From Git Bash call them with
`MSYS_NO_PATHCONV=1 cmd.exe /c .\scripts\x.cmd`: otherwise `/c` becomes a path and an interactive cmd opens.
A new `.ps1` also needs: its name in the two lists of "help of every script" and a `Describe` in `tests/ps/Scripts.Tests.ps1` (the
scripts it runs are faked by functions named like their full path), its `J2D-SCR-nnn` rows in `error_codes.py`, a row in the README table.

A release holds the workbook only (AL, 2026-11-12: no zip of the code); the "Source code" archives that GitHub adds to every
release (they cannot be removed) leave out the `export-ignore` paths of `.gitattributes` (CLAUDE.md, `archive/`,
`.claude/`, `.venv/`); `tests/unit/test_release_contents.py` pins the list. `-Publish` is outward-facing: ask AL first.
Releases: the version is `vX.Y.Z` only (`v0.0.0-dryrun` is refused) and the number is AL's choice. The dry run
(`release.ps1 -Version v0.0.0`) needs a clean tree, builds with Excel and leaves `out\release\<version>`: delete it afterwards
(only that dry-run folder: a real release keeps its `out\release\<version>`, whose `notes.md` holds the SHA-256 and MD5). The release body
starts with `(commit <40 hex>)`: take the MD5 to note here from the line ending `bytes)` or from `md5sum out\release\<version>\jira2dashboard.xlsm`, never "the first 32 hex".
`-Publish` also needs HEAD == origin/main and `gh` logged in. Only `v0.5.0` (2026-11-13, commit e01327b), `v0.5.1` (d67e42f), `v0.5.2` (f541e99), `v0.6.0` (57b0d79), `v0.7.0` (2026-11-14, commit cc01848) and `v0.7.1` (2026-11-14, commit db81283) and `v0.8.0` (2026-11-15, commit 5ea633c) and `v0.8.1` (2026-11-15, commit d2035ab: the buttons stop with J2D-SYS-003 when the workbook is locked) and `v0.9.0` (2026-11-16, commit 209f229: Feature Predictability from the PI field history, IterationFeatures and IterationCharts, scenario art-metrics) and `v0.9.1` (2026-11-17, commit 4722b9d: a Program-only ART needs no Team project, text of the Manual; MD5 bbb9220ad709404142953179307b2443) and `v0.10.0` (2026-11-17, commit 41269f9: the value above the Velocity bars and notes under the charts with gaps, the chart Cycle time by status and the table PIStatusTime; MD5 e3612463f3aee8235f7ea64d572d16fd) and `v0.11.0` (2026-11-18, commit 9f674cd: the sheets FieldMapping, Order and Type are gone, Team optional, Propose values fills IsEnd, the issue-type lists and BlocksLinkType, the table VelocityCheck; MD5 bfe9145574abe7fd750a60d5c3030d7b) and `v0.12.0` (2026-11-18, commit 2d61724: the ART counts the Features only, ART (other types), MetricItems and Reconciliation; MD5 08d5fce1cd1975ef28a1eebea8ff170f) and `v0.13.0` (2026-11-19, commit abb84bf: statuses, workflows and moves read from Jira (Jira-first rules, the sheet Workflows holds exceptions), ItemFlow.JiraWorkflow and DC28 by the Jira workflow, the Cycle time by status stack ordered by the Jira flow, Propose no longer copies statuses; MD5 274cc63ede1d7adb0d179af59ab9a286) and `v0.14.0` (2026-11-23, commit f70df7a: why a flow metric is not computable, told
in the log, the health fields, the Reconciliation and the chart notes, no start status or no Wait status first; the date guards of the audit; MD5
c74fbacb63b9cc609cc5b2f3e4ce24b5) and `v0.15.0` (2026-11-23, commit a8db023: the three buttons 1. Extract from Jira, 2. Compute the metrics, 3. Draw the
charts; Compute makes the status rules again without reading Jira; MD5 4e1d7b42616e685281bde9773d11480b) and `v0.16.0` (2026-11-24, commit 80c8f37: the four
accounting tables FlowTimeCheck, FlowEfficiencyCheck, FlowLoadCheck, FlowDistCheck, 18 controls, the work-type bucket ignores case; MD5
5a3287c4cda38b0ca3f3ea1f51cf28cb) and `v0.17.0` (2026-11-24, commit 57820b0: the report of 236 figures, 20 controls with CarryOver and the Flow Load list,
VelocityCheck.Reason, MetricItems with the days and bucket of each ticket and the in-process list; MD5 9071979105af953fbfdfdf5270540671) exist on GitHub: AL removed the
releases and the tags `v0.1.0` to `v0.4.0` (2026-11-13, they were 2026-11-06, 878807e, 31c9574, f0bbf32). Their numbers are never used again; the next one is above `v0.17.0`.
A clone may keep old tags (`git fetch --tags`): `version.py` stamps the commit, never `git describe`.
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
never cleaned either; `out\inter` holds the rest (test workbooks, exports, `.coverage`, `.pytest_cache`, `logs`),
emptied by `build` / `vba-tests` except the caches and `logs` (`src/jira2dashboard/clean.py`). Tests write to `INTER_DIR`, with one
exception (AL, 2026-11-17): the chart PNG that the Excel scenarios export to be looked at go to `out\manual_test\charts\<scenario>\`
(`paths.CHARTS_DIR`, `flow/sheets.export_charts_to_look_at`; `flow` = the 9 + 5 charts of `flow_excel.feature`, `art_metrics` = the
scenario art-metrics), so `build` and `vba-tests` do not erase them. Each chart is exported twice: `<name>.png` and `<name>.svg` (Excel
writes the SVG itself: `Chart.Export(path, "SVG")`, a vector, so a zoom stays sharp). Each run empties the folder first (no image of a
removed chart) and writes AFTER the export `_exported.txt` (UTC time, commit, image list) and `index.html`, a dashboard of the SVG with a
click-to-zoom (wheel, drag, Esc; no script or font from outside: `chart_images.index_html`); `out\manual_test\charts\index.html` links the
scenarios. No note, or an old time, = the last run did not reach the export. Open the page by double-click. Rerun the scenario to
refresh: `test-excel tests/steps/test_flow_excel.py -k "art_metrics or png"` (the second test is the `flow` one).

Before finishing a change run the full pass: `coverage`, `vba-tests`, `test-excel` (flow module with `-n 4`, then the rest; both `-k "not pipeline"`), `test-scripts`, `build`
(skill `/full-pass`; `test-network` and the `@pipeline` scenarios only when the build pipeline changes: they replay the VBA suite that `vba-tests` runs; a pass takes about 6 min). The Claude config is versioned in `.claude/` (like this file) but left out of the release
(`export-ignore`); only `.claude/settings.local.json`, the personal settings, is gitignored. Hooks (`.claude/hooks`):
`bas_guard.py` runs the VBA guard tests after an edit of `vba/*.bas`, and after a Bash/PowerShell command that changed a
module (sizes and times kept in `%TEMP%\jira2dashboard\bas_guard.json`); `excel_kill_guard.py` asks before a command that
stops Excel by its name (`taskkill /IM`, `Stop-Process -Name`). Skills: `full-pass`, `add-data-check`, `read-logs` (summary
of the logs: `summarize_logs.py`), `release` (user only), `excel-hygiene` (Claude only: count `EXCEL.EXE` before and after a run, stop
only the recovery or VBA-dialog Excel the run started, by PID), `vba-mutation` (user only: `mutate.py apply|restore|status`, the mutation proof of a
`T_` test; run `status` before a commit). Agents (`.claude/agents`, read-only): `vba-reviewer` (a `.bas`
change against the VBA traps, before `vba-tests`), `parity-checker` (VBA vs Python path, formulas vs `flow/`, modLog vs
python.log). `tests/unit/test_claude_hooks.py` and `test_claude_skills.py` test the scripts (skipped without `.claude`).

## Architecture (big picture)

- **Sheet names say where the data comes from (AL, 2026-11-13).** `_from_user` = typed by the person (Config, Projects, Iterations, PIs, Workflows,
  Transitions), `_from_Jira` = read from the Jira API (Issues, StatusIntervals, IssueLinks, the `Jira...` proposals), `_from_excel` = calculated,
  drawn or written by the workbook (Charts, Metrics, ItemFlow, Dashboard, Manual, Steps, the plan sheets...). The code keeps the LOGICAL names (`Config`,
  `Issues`; the Excel tables too: `Issues[Key]`) and converts at the edge: `sheet_names.of(name)` in Python, `modText.SheetName(name)` in VBA (both idempotent,
  an unknown name such as a test's scratch sheet comes back unchanged; `LogicalSheet` goes back). The three lists must be the same on both sides
  (`tests/unit/test_sheet_names.py`); a new sheet is added to both, then to the scenario "chaque feuille ... suffixe". A message may name a sheet by the
  start of its name ("the Config sheet"). Never name a VBA variable or parameter like a procedure (VBA ignores case: `sheetName` hid `SheetName`,
  `readWorkflows` hid `ReadWorkflows`; only `vba-tests` shows it, as COMPILE_FAILED).
- **Python generates, Excel computes.** `src/jira2dashboard/build_workbook.py` writes the parameter sheets, the empty raw
  sheets and a *plan*, injects the `vba/*.bas` modules and compiles them. VBA writes only raw data (`Issues`,
  `StatusIntervals`); **every figure is an Excel formula** in a table (`ItemFlow`, `DailyLoad`, `Metrics`, `DataChecks`).
- **Formulas and charts are data.** `flow/formulas.py` and `flow/dashboard_spec.py` hold the formula text and chart
  descriptions (pure, tested without Excel). `flow/blueprint.py` writes them into the sheets `FlowFormulas`
  (Table|Column|Formula; empty Formula = a values column), `ChartSpec` and `ChecksSpec`. After each extraction VBA
  `modFlowBuild.BuildFlowSheetsIn` rebuilds all tables, formulas and charts from those sheets, sized to the new data.
  Add a metric or check by editing the plan, not VBA.
- **Feature Predictability (AL's decision, 2026-11-12).** A variant of Flow Predictability, whose SAFe definition (business
  value of the PI objectives, `docs/` section 2.2) stays as a chart, not computable. **Since 2026-11-16 (AL) a PI is judged on
  what was ENGAGED at its start, read in the changelog of the PI field, not on the PI field as it is now** (a Feature deferred
  later, PI emptied or moved, would leave the denominator; one added later would inflate it). Per PI of the sheet `PIs`, for the
  ART: planned = the Features (`FeatureTypes`) that carry the PI at the start of its first iteration (the baseline), done =
  released in an iteration of the PI (`PredictabilityMode` A, default: among the planned; B: also the added); a cancelled Feature
  stays planned, not done; nothing planned = a gap. Both come from the table `IterationFeatures` (one row per iteration x PI that
  holds its start x team: the ART, then the teams of the Features = their Team field; Engaged, Added, Deferred, Released,
  ReleasedPlanned, ReleasedAdded, CarryOver, with Engaged(next) = CarryOver), itself counted by COUNTIFS on `PIIntervals` (raw,
  written from the changelog by `modFlow.PiIntervalRows`: one row per stretch in which an issue carried a PI, touching stretches of
  the same PI merged). Reference `flow/pi_history.py`, `flow/iteration_features.py`, `flow/metrics.pi_predictability`; a PI with no
  iteration has no baseline (a gap). Existing metrics keep Features under the team ART; `FeatureTeam` is for this table only. Then
  the valid PIs (`PI#XX-YYYY`) that Features carry but `PIs` lacks, by year then number, not computable, counted by the PI field as
  it is now (AL, 2026-11-12; `modFlowBuild.FeaturePis`). The manual-test copy declares both PIs of the generator (its Features are
  created before the PIs start, in the generator, so that they are engaged). Table `PIMetrics` (`formulas.PIMETRICS_COLUMNS`; keys
  `modFlowBuild.PiNames`), chart point `dashboard_spec.pi_point`, the 5 charts of the sheet `IterationCharts`
  (`dashboard_spec.ITERATION_CHARTS`, `ChartSpec` column `Sheet`) and their SVG in `charts_svg`. Known and left: a Feature done in
  an iteration where it left the PI and came back counts twice in the workbook (`iteration_features._flags`; pinned by no scenario).
  Names are not escaped: a team or a PI name with `*`, `?`, `~` or a leading `=`, `<`, `>` is read as a wildcard or an operator by
  `COUNTIFS`, which also ignores case and trims only the ends by `TRIM` (Excel) / `TrimWhite` (VBA) / `.strip()` (Python) differently
  for inner double spaces and non-breaking spaces; a team named `(none)` would match the Features without a team (`PIIntervals.FeatureTeam`).
  The PI figures are sums of iteration rows: PIs and iterations that do not tile each other (an iteration that starts outside every PI,
  a Feature done after the last iteration of its PI) leave figures out; the PIs of the manual-test copy share one window, so each
  iteration has a row per PI.
- **Scenario `art-metrics` (AL, 2026-11-16).** `mock-jira --scenario art-metrics` / `mock-workbook --scenario art-metrics`
  (`mockjira/art_metrics.py`, `scenarios.py`): 25 Features of the Program project over 4 PIs (`PI-1`..`PI-4`, 126 days, 6 iterations of
  21 days, no gap), workflow `program-workflow` (Jira side; the sheets keep `Feature`), PI field history and the Iteration field in the
  changelog/fields (`FieldChange`), 3 teams in the Team field, deferrals (PI emptied or moved to `PI-5`) and ft-25 added during PI-4. Every
  expected number is typed BY HAND in `tests/art_oracle.py` (oracle L0): unit tests, `art_metrics.feature` (Python path) and
  `art_metrics_excel.feature` (the real workbook, also against the reference cell by cell) compare to it, never only to `flow/`.
  **A scenario is a SEPARATE INSTANCE with only its own data (AL, 2026-11-17).** Its parameter sheets are built from its own text
  (`art_metrics.sheets()`): Projects = the Program `TTT` alone (no Team project: the legends of its charts show `ART`, never `AAA`...),
  Workflows and Transitions = the Feature workflow alone, its own Iterations and PIs; only Config (the product's list of
  parameters and columns) come from `params/`. Its `FieldIds` keep only the fields it uses (PI, Team, Work type, Iteration; the others are
  `None`: not in `/field`, not in the issues, empty in its Config). The product rule was relaxed for it (`params/model.shape_problems`):
  the Team workflow is needed only when a Team project exists, and a Team project is no longer required. Its fake Jira listens on its own
  port (`Scenario.port` = 8101; the generic one keeps 8099, both can run), its copy is `out\manual_test\art-metrics\jira2dashboard4art-metrics.xlsm`
  (own `logs\`, own Rovo CSV; `mock_workbook.copy_path`, `fill_workbook` is the one writer of the parameter sheets, also used by the Excel
  steps). **Run it alone:** `scripts\manual-test.ps1 -Scenario art-metrics` (fake Jira + copy + Excel, you click; `-Python` = no macro) or
  `-Scenario art-metrics -Run` (nothing to click: `mock-workbook --scenario art-metrics --run` serves it in-process, extracts with the
  macros, exports the charts to `out\manual_test\charts\art_metrics\`, saves the copy with its results, closes Excel, opens the page);
  `start-mock.ps1 -Scenario NAME` is the fake Jira alone. The page `index.html` of `charts\art_metrics` opens with the report **expected
  (worked out by hand) / obtained (workbook)**: 236 figures (PI predictability, IterationFeatures of the ART and of each team, Velocity,
  Flow Time, Flow Efficiency, Cycle time by status, the four accounting tables FlowTimeCheck... of the ART: `CHECK_TABLES`), the sections that differ open and the lines that differ first (`tests/art_report.py`, pure `sections` /
  `render`, `report_for_workbook` reads the tables through COM; the oracle stays in `tests/art_oracle.py`). The Excel test writes it
  and asserts "all 236 are OK"; `--run` loads it from `tests/` by path (`mock_workbook.report_builder`, none outside a checkout). `ps1` ports: `$ScenarioPort` must equal `Scenario.port` (a test pins it).
- **Cycle time by status (AL, 2026-11-17).** The chart `pi_status_time` (stacked 100 % columns, PI on the X axis, one colour per status) and the
  table `PIStatusTime` (sheet `PIStatusTime_from_excel`: PI | Status | Start | End | Items | TotalDays | AverageDays | Share) show where the cycle
  time of the Features goes, to find the statuses that take most of it. A PI = the Features RELEASED in it (DoneUtc in [Start, End) of the sheet
  PIs, not excluded from flow time); the statuses = the Active and Wait statuses of the Feature workflow in the ROW order of the sheet
  Workflows (the sheet has no `Order` column since 2026-11-18: its rows are the normal flow; Unclassified is left out); TotalDays = the time those Features spent in the status from their
  StartUtc (included) to their DoneUtc (excluded), the window of `ItemFlow.ActiveDays/WaitDays` (so the totals of a PI add up to the Active + Wait
  days of the same Features: a test asserts it); Average = Total / Items; Share = Total / the total of the statuses of the PI. The helper columns
  of `StatusIntervals` (`ItemRow`, `ItemStart`, `ItemDone`, `ItemExcluded`, `InCycle`) feed it. Reference `flow/pi_status_time.py`, hand-typed oracle
  `tests/art_oracle.PI_ITEMS/STATUS_AVERAGE_DAYS/STATUS_SHARE` (Items 6, 5, 3, 5: ft-25 is released in PI-4; 1 day per status, 2 for the long one = 25 %),
  hard cases (a Feature crossing a PI, a status visited twice, a backlog return, reopened, half-open window) in `tests/unit/test_pi_status_time.py`.
  VBA: `modFlowBuild.StatusNames` (the series of the chart: its plan `SeriesNames` is the token `{statuses}`, read from Workflows at the build),
  `SelectorColumn` (the selector list sits right of the widest ChartData block: teams, Feature teams or statuses), `WritePiStatusKeys`.
  Python path: sheet `PIStatusTime`, SVG chart. The report of the scenario art-metrics has a 6th section.
- **VelocityCheck (AL, 2026-11-18).** On the client PC the Velocity of the ART read 4 / 10 / 60 where Jira's `status CHANGED TO "Terminé" DURING` read +8 in
  the last sprint. The table `VelocityCheck` (sheet `VelocityCheck_from_excel`, the keys of `Metrics`, reference `flow/velocity_check.py`) counts the Velocity again
  from the intervals of `StatusIntervals`: `Velocity`, `VelocityFeatures` (types of `FeatureTypes`: **the team ART is every issue of the Program project, `Velocity`
  has no type filter**), `CreatedDone` (done because CREATED in an IsEnd status: no change in the changelog), `ChangedToEnd` / `ChangedToEndFeatures` (changes into
  IsEnd, one per change), `ReopenedLater`, `Gap` = Velocity - CreatedDone - (ChangedToEnd - ReopenedLater), always 0. Helper `StatusIntervals.ItemIsFeature`.
  `Reason` (AL, 2026-11-24; a property of the Python row, `velocity_reason`, TEXTJOIN in Excel): why the Velocity differs from Jira's count of the changes into an
  end status (created done, closed then reopened, other types of the Program), or that nothing was done; "" when it is that count.
  Hand-typed cases in `tests/unit/test_velocity_check.py`; `flow_excel.feature` (cell by cell) and `art_metrics_excel.feature` (= the oracle's Velocity).
- **The ART counts the Features only (AL, 2026-11-18).** `ItemFlow.Team` of a Program ticket is `ART` when its type is in `Config.FeatureTypes`, else `ART (other types)`
  (`flow/model.OTHER_TYPES`, `program_team`; formula `formulas.ITEMFLOW_COLUMNS` Team; Python `flow_adapter.to_flow_issues` reads FeatureTypes from the sheet Config,
  `mapping.flow_issue/extracted(feature_types)` from the settings). No metric counts the other team (the teams list is ART + the Team projects), they stay in ItemFlow
  (findable by key), the data checks and `VelocityCheck.OtherTypes`; `PIStatusTime` reads `Team = ART`, not the workflow; the Python `fill_teams` leaves the other team out.
  `modHealth.NotFeatures` says `Not Features: N of M` and, when the ART has no Feature, that `FeatureTypes "X" matches no ticket` (the client's type is "Fonctionnalité",
  the delivered default stays `Feature`). Scenario art-metrics holds 3 Program sub-tasks (TTT-26..28) that no ART figure counts. Sub-tasks of Team projects are still counted.
- **Jira-first status rules (AL, 2026-11-18).** The metrics read the statuses from Jira; the sheet `Workflows_from_user` holds EXCEPTIONS (an empty cell takes Jira's value, an
  empty `Workflow` touches both kinds, a row of the kind beats a row of both). Per status name (`flow/status_rules.py`, VBA twin `modParams.EffectiveRules`, same cases in
  `T_StatusRules`): Category from the Jira category (new -> Backlog, indeterminate -> Active, done -> Done), IsEnd = every Done status (exceptions: Removed / Wait / Backlog are no end),
  IsStart = first in-progress status from the creation status in the transitions graph (`first_in_progress`, computed in `JiraWorkflows` at the Jira read and by
  `modJiraConfig.ReadWorkflowRules` at an extraction that has no such sheet), IsFree = GLOBAL transition. Sources, in order: the exception, the catalog of every status of the instance
  (`ReadStatusCatalog`, GET `statuses/search`, sheet `JiraStatusCatalog`, tolerant of a refusal), the category Jira gives each ticket's CURRENT status; a status that none knows stays
  `Unclassified` (a retired workflow's status in a history). `modExtract.RulesInit` merges them once per run and writes the sheet `StatusRules_from_excel` (Workflow | Status | Category |
  IsStart | IsEnd | IsFree | Source Jira/user); `StatusIntervals.Category` still comes from it at extraction time, and the formulas `IsStart/IsEnd/BadMove` read it (letters A..F as the old
  Workflows layout). Flow Efficiency is `not computable` without a status declared Wait (`Rules.waits`, formula `COUNTIF(StatusRules...,"Wait")`, `dashboard_spec.EFFICIENCY_REASON`).
  The Python path asks the same catalog (`JiraClient.status_catalog`) but not the workflows (no start / free from Jira there: the exceptions give them). The delivered
  `params/workflows.md` stays as a complete example. **The workflow of each ticket (AL, 2026-11-18)**: the extraction without a sheet `JiraWorkflows` reads the projects and the
  workflow previews itself (`modJiraConfig.ReadWorkflowRules(..., wb)`; Python `standalone/workflows.read_workflows`) and WRITES `JiraIssueTypes`, `JiraWorkflows`, `JiraTransitions` (created
  empty at build: formulas read them); reference `flow/jira_workflows.JiraWorkflows` (workflow of (project, type), creation status, moves, starts, free). `ItemFlow.JiraWorkflow` (last column)
  = that workflow of the ticket, `StatusIntervals.JiraWorkflow` its copy. **`ItemFlow.Workflow` stays Team/Feature (the kind the rules are keyed by), it is NOT the Jira workflow.**
  **DC28** (`BadMove`): a move is allowed by the Jira workflow of the ticket OR by `Transitions_from_user` (kind), free statuses pass; a ticket whose moves nobody knows (no row of the sheet for its
  kind, no Jira workflow read) is NOT judged (`checks.Moves.permits`, scenario art-metrics: the 3 sub-tasks that skip steps are the only flags, hand-typed `art_oracle.DC28_*`).
  **The stack of the chart Cycle time by status** (`PIStatusTime`) no longer reads Workflows rows: the Active/Wait statuses the Features (Program project, FeatureTypes) visited, ordered by
  depth in the Jira workflow of the Feature types (`status_rules.flow_depths`, `JiraWorkflows.feature_depths`), then earliest `Seq`, then name (`pi_status_time.feature_statuses`,
  VBA `modFlowBuild.StatusNames(wb)` reading the RAW sheets). **The scenario art-metrics keeps Jira's full definition in `Dataset.jira_sheets` and only the 3 Wait exceptions in its
  `sheets`** (`flow_adapter.jira_rules/jira_workflows` give the references): a test that builds `Rules` from `ds.sheets` alone sees no start. Not done yet: the delivered `params/workflows.md` / `transitions.md` still hold the complete example. The Python path writes the Jira sheets it read (JiraStatusCatalog, JiraIssueTypes, JiraWorkflows, JiraTransitions in the VBA layout: `jira_workflows.sheet_rows` / `from_sheets`, headers pinned against `modJiraConfig.bas`; the Python reading keeps no status id, Category, IsEnd or Project: empty cells) and StatusRules (`Rules.entries` keep the source).
- **MetricItems and Reconciliation (AL, 2026-11-18, "c'est de la comptabilite").** Every figure can be traced back to its tickets. `ItemFlow` gets the trace columns
  `DoneIteration` (+ `DoneIterationStart/End`, found in the key columns of `Metrics`: no circular reference), `DonePI` (from `PIMetrics`), `Account` (the partition of the
  population: done in an iteration / done outside the iterations / left out: not a Feature / removed / open), `WhyNotInFlowTime` (Reopened, No valid start),
  `BoundaryHours` (distance of DoneUtc to the nearest iteration boundary); reference `flow/trace.py`. The sheet `MetricItems_from_excel` holds three lists of SPILLED
  formulas side by side (columns A.. Done in an iteration: Iteration, Start, End, Team, Key, Type, DoneUtc, FlowTimeCounted, FlowTimeDays, WhyNotInFlowTime,
  ActiveDays, WaitDays, UnclassifiedDays, WorkTypeBucket (2026-11-24); P.. done but not counted with the reason; W.. cut-off within 24 h; AC.. in process at the end of
  each iteration, the tickets behind FlowLoadEnd and DistWip: ONE spilled formula in its first column (the other columns have no formula: BuildMetricItems skips an empty
  one), an iterations x candidates matrix (mode A the intervals of StatusIntervals, mode B the tickets of ItemFlow, the rule of `FlowLoadCheck.AtEnd`), the pairs of a team
  whose FlowLoadCheck.Reason is not "" left out, flattened with INDEX and CHOOSE (not TOCOL/HSTACK/CHOOSEROWS); reference `metric_checks.in_process_at_end`; LET names
  `r`/`c` are refused by Excel (R1C1), use `rw`/`cl`; measured 0.043 s for the MetricItems step on 890 intervals, 0.040 s before), `modFlowBuild.BuildMetricItems`,
  python `trace.metric_items`; the rows of an iteration and team are its Velocity. **A reader of a list takes its column from `formulas.metricitems_column` /
  `metricitems_cell`** (the ledger of the Reconciliation, its cut-off control, `metric_blocks` of the Excel steps): a column added to a block moves the next blocks.
  The sheet `Reconciliation_from_excel` has the table `Reconciliation` (20 controls since 2026-11-24: + a start status per workflow kind with tickets, a Wait status, the Gap of the
  four flow check tables, and CarryOver of IterationFeatures = the Engaged of the next iteration of the same PI and team: `iteration_features.carry_over_gap`, an
  XLOOKUP of End|PI|Team in Start|PI|Team, a gap or another PI not judged; NOT Engaged + Added - Deferred - Released = CarryOver, which is the formula of the column
  itself and could never fail; an Excel step types one Engaged over by hand and sees the control find 1; Flow Load at the end of all iterations = the rows of the
  in-process list of MetricItems (`SUM(Metrics[FlowLoadEnd])` against the rows of its spill); Control | Assertion | Expected | Found | Difference | Result OK/REVIEW, in the order
  of `reconciliation.CONTROLS`; Expected and Found are a formula of their own per row, written BEFORE the table exists: **a formula typed in one cell of a table column fills
  the whole column**, which made all the rows equal), the census of the types (spilled from H1) and the parameters in force; reference `flow/reconciliation.py`;
  `modHealth.ReconciliationFindings` says `Reconciliation: n of m controls to review` and counts the `REVIEW` rows only (the cut-off control is `INFO`: `reconciliation.INFO_ASSERTIONS`);
  the statuses control counts the tickets that are not in "ART (other types)" and the unclassified statuses are listed by name from column O (`reconciliation.unclassified`). The art-metrics scenario types every list and control by hand.
- **One accounting table per flow metric (AL, 2026-11-22: "the same accountability checks as for predictability").** `FlowTimeCheck`, `FlowEfficiencyCheck`, `FlowLoadCheck`,
  `FlowDistCheck` (sheets `..._from_excel`; `FlowDistributionCheck_from_excel` would pass the 31 characters of an Excel sheet name), the keys of `Metrics`, reference
  `flow/metric_checks.py`, the same sheets in the Python path. Each holds a funnel of counts, a `Gap` between TWO INDEPENDENT paths (always 0: a recount of the same ItemFlow
  column would be 0 by construction and no mutation could fail it) and a `Reason` that is "" exactly where the metric is computable (the Excel scenarios assert it against
  `Metrics`). Flow Time: Done recounted from the intervals minus the left-out items (reopened / no start status in the rules / no start after the backlog) vs `FlowTimeCount`.
  Flow Efficiency: Active and Wait days summed from `StatusIntervals` (`InCycle`) vs `ItemFlow`. Flow Load: the load at the end by SUMPRODUCT (mode A over `StatusIntervals`,
  B over `ItemFlow`) vs the COUNTIFS of `DailyLoad`; reasons no start status, every day after AsOf. Flow Distribution: COUNTIFS on the RAW `ItemFlow.WorkType` vs
  `DistDone* x Velocity` (in Python both sides go through `work_type_bucket`: there the Gap is 0 by construction, only Excel's is a real check), a `Note` when every done item is Unmapped (WorkTypeField empty, or no known work type). **The work-type bucket ignores case and keeps the canonical
  name on both sides** (`model.work_type_bucket`, nested IF of `WorkTypeBucket`: Excel's `=` counted "feature" as Feature, Python did not; a number is Unmapped, never a
  crash; the generated data holds "defect" and tickets with no work type, so the Excel scenarios see both). Hand-typed oracle in
  `tests/art_oracle.py` (`FLOW_*_CHECK`, with and without a start status); one Reconciliation control "Gap = 0" per table.
- **Python reference as oracle.** `flow/` (`intervals.py`, `itemflow.py`, `load.py`, `metrics.py`, `checks.py`) is a pure
  Python implementation of the same metrics and data checks. The scenarios in `tests/features/flow_excel.feature` run the
  real delivered workbook against the fake Jira and compare Excel with this reference **cell by cell** (Flow Load modes A
  and B). Keep the two sides in step: a formula change needs the reference change and vice versa.
- **Parameters live in the workbook, never in code.** `params/*.md` (config, projects, workflows, transitions,
  iterations) are read and validated by `src/jira2dashboard/params/`, written into sheets, and read by VBA
  (`modParams`). Status names, team names and extract columns must not appear as literals in VBA (a guard test enforces it).
  `Config` columns: Parameter | Value | Example | Mandatory | Description (formulas read A and B; the rest by header). An empty Mandatory=True value
  stops the buttons before any call to Jira (`modParams.RequireMandatory`, error 608).
- **Dashboard and Charts.** Top left, one under the other (AL's order): the progress bar `A2`, the buttons Prompt for Rovo
  `A3`, Load settings CSV `A5`, Read the Jira configuration `A7`, Propose values from Jira `A9`, then the three steps `1. Extract from Jira` `A11`,
  `2. Compute the metrics` `A13`, `3. Draw the charts` `A15`, Cancel `A17` (cells and captions in `dashboard_spec`, the captions also in `modActions`: a test). A
  final status is `Title | Name: value | ...`: `A20` (name `dashStatus`) shows the title and `A21:B40` one row per field (`modText.StatusParts`, `modActions.ShowProgress`; `ProgressSnapshot` rebuilds the line from the
  cells for the tests). The 10 charts, their notes and the team selector `B2` (name `dashTeam`) are on the sheet `Charts`, made
  at build after the Dashboard (`CHARTS_SHEET` of `modFlowBuild`/`modActions` = `dashboard_spec.CHARTS_SHEET`, a test pins it). The 5
  charts of the Features per iteration are on the sheet `IterationCharts` (right after `Charts`; `ChartSpec` has a column `Sheet`, the
  rows of `IterationCharts` come BEFORE the daily curve: its block of `ChartData` is the longest, so it stays last; the helper
  `chart_blocks` of the Excel steps follows that order). **An iteration where nothing was done** (AL, 2026-11-17: every Feature deferred,
  say) has Velocity 0 (no bar) and Flow Time / Flow Efficiency "not computable" (a gap, never 0): `ChartSpec.labels` (plan column `Labels`,
  text "yes", `modFlowBuild` FLB49) writes the value above the bars of the first series (the ART) on the Velocity chart only, and the note
  under the chart says why (`ChartSpec.zeros`: "N of M points are 0 - no item was done (no bar)"; `reason=NO_ITEM_DONE` under Flow Time,
  p85 and Flow Efficiency; `charts_svg.gaps_note` is the Python twin). `-Run` and `manual-test` use `out\release`: run `build` first after a VBA change.
- **User manual.** `src/jira2dashboard/manual.py` (pure text) is written by `flow/sheets.write_manual` at the delivery
  build: the first steps on the Dashboard from T1, and the sheet `Manual`. Update it when a button, sheet or step changes.
- **Extraction (VBA).** `modHttp` (WinHTTP, proxy with Windows auth, retries; `modProxy` finds the proxy of the Windows
  settings when `ProxyServer` is empty: WPAD, PAC or fixed, via WinHTTP API calls on byte buffers, not VBA Types, which
  crashed 64-bit Excel) -> `modJira` (search, bulk changelog) ->
  `modExtract` -> `modFlow` (status intervals) ; `modActions` holds the buttons. The previous
  results are cleared only after Jira has answered. **Three steps (AL, 2026-11-22: correct and compute again without reading Jira):**
  `BtnExtract` (the slices; ends with `Extracted | Issues | Status intervals | Changelog rows | Next`, the raw sheets only, the old tables and charts
  deleted), `BtnComputeMetrics` (`ComputeMetricsCore`: the compute parameters, `ValidateBuildInputs`, `modExtract.RefreshRules`, `BuildTablesIn`; no call to
  Jira) and `BtnDrawCharts` (`BuildChartsIn`, error 712 `J2D-FLOW-007` without the tables); `ExtractAndCompute` (tests, `--run`) is the three in one go.
  Each button checks its own mandatory parameters (`EXTRACT_PARAMETERS` in modExtract, `COMPUTE_PARAMETERS` in modActions): a client may extract before
  filling Iterations and PIs. **`RefreshRules`** makes the status rules again from what the extraction kept (`JiraStatusCatalog`, now written at every
  extraction even empty; `JiraWorkflows`; Issues.Status/StatusCategory) under the exceptions of Workflows, then rewrites StatusIntervals `Workflow`,
  `Category` and the open `DurationDays` in place: a change of Workflows, of a Role of Projects or of AsOf needs no extraction; the field ids, Filter,
  the project list and BlocksLinkType do. It is NOT in `BuildFlowSheetsIn` (`T_FlowBuild` writes StatusRules by hand); `RulesInit` stays at the extraction and
  every flow scenario computes right after extracting, so a difference between the two shows against the reference. Measured before the buttons: 15 rounds
  of tables then charts in one Excel, no crash (the open defect hit a second build at `flow sheets: charts`). `modLog` writes the local log in the folder `logs` next to the workbook (it runs from a folder not known in advance; `%LOCALAPPDATA%\jira2dashboard\logs` is the fallback):
  see **Logs** below; never pass it a token, URL, JQL or Jira answer (`test_vba_log_guard.py` checks).
- **Python path without VBA (`standalone/`, `run.ps1 extract`).** A second consumer of the flow reference: it reads
  the parameter sheets of the filled workbook as a zip (`xlsx_read`, by header name, shape rules shared with
  params/*.md through `params.model.shape_problems`, Config rules of the VBA in `params_book`), calls Jira with the
  WinHTTP COM object (`http_winhttp`: same proxy, Windows auth and retry rules as `modHttp`; body read from
  ResponseBody as UTF-8; injectable sleep) after `proxy.py` (ctypes port of `modProxy`), maps the JSON like
  `modExtract`/`modFlow` (`mapping`, oracle: `test_mapping` against `flow_adapter`), computes with `flow/`
  (`results`, same column names as `formulas.TABLES`), draws the 10 charts as SVG (`charts_svg`) and writes values +
  pictures with zipfile (`xlsx_write`: SVG with a blank PNG fallback, Excel 365 shows msoGraphic). Keep it in step
  with the VBA when an extraction rule changes. Its log (`logs\python.log` next to the workbook, as the VBA log) is complete on
  purpose (AL's choice): status, call, Jira's message, tracebacks; never the token nor the Authorization header
  (`Scrubbed` formatter; `test_extract_main` checks).
- **Changelog sheet (AL, 2026-11-13).** `Changelog_from_Jira`: one row per changelog item of the tickets read, every field (`modFlow.ChangelogRows`, python
  `mapping.changelog_rows`; columns `formulas.CHANGELOG_RAW` = `modFlow.ChangelogHeader`, a test compares them): Key, IssueId, ChangeId, CreatedUtc (the date read, empty
  when it cannot be) and CreatedText (as Jira sent it: the unreadable ones show), FieldId, Field, FieldType, From, FromString, To, ToString. **No author** (AL's choice),
  texts cut at 255 characters (AL's choice), no value for `assignee`, `reporter`, `creator` and every field whose schema type or items is `user` (one call to
  `field` at the start of the changelogs: `modJiraConfig.UserFieldIds`, `JiraClient.user_fields`; a failed call gives none), `[~accountid:...]` becomes `[user]`
  (a name is personal data; a name written in clear in a description stays: only leaving out the long texts removes it). Written last by `WriteRawSheetsCore`, cleared by
  `ClearFlowSheets`; the fake Jira only sends status changes and key moves, so its rows are `len(changes) + len(key_changes)`.
- **Dates and times (audit of 2026-11-13).** All dates are UTC. `modText.ParseJiraDate` and `standalone/mapping.parse_jira_date` must agree (offsets, `Z`, a space for `T`,
  a number or digit text of 1e11..1e14 is epoch milliseconds, 1e9..1e11 seconds: the real `changelog/bulkfetch` sends `created` as epoch ms (1761134538011) although the OpenAPI says
  `date-time`, found 2026-11-14 when every status change was left out; PostHog PR #123173 saw the same). **The date of a history is found by its VALUE, never by its label**:
  `modJira.FindDateKey` / `mapping.find_date_key` read the first 20 histories and take the key that holds a plausible date (2000 to tomorrow; `modText.DateKind`,
  `IsPlausibleDate`) in every one: `created`, else a name with `creat`/`date`/`time`, else the first; one `INFO changelog date key: <key> (<kind>)` line in the log; the fake
  Jira can rename the key and change the format (`MockJira(changelog_date=("createdAt", "iso"))`). A date nested in an object is not read (falls back to `created`
  with the `J2D-DATA-003` warning); `modParams.ParseMoment` reads an hour as the sheet formula does (`H:MM`, not dropped); an Iterations date before 2000 is refused
  (error 701: a typo such as 1026 would write 365 000 days per team); two changes in one second are ordered by their id as a NUMBER (`model.change_order`), as the
  VBA keeps Jira's order; the iterations span is the earliest start to the latest end (`results.iteration_span`). Known and left: the logs use local time (VBA and
  python.log, unlabelled) while `Run` and `Errors` use UTC; an unreadable `created` crashes the Python path and gives a huge first duration in VBA (Jira always sends it);
  the fake Jira always sends `+0200` (no DST mix). **Guards of the audit of 2026-11-23** (same rules in VBA and Python): the time of a Jira date is read by its
  separators (`...T09:12+0200` read `02` as seconds); dates of Iterations, PIs and AsOf are from 2000 to 2099 (`20261005` typed as a number was the year 57 000),
  Iterations and PIs in whole days, the PIs checked at the build (error 710, `J2D-CFG-013`: a French text `05/10/2026` was read by `DATEVALUE` and not by the VBA),
  an offset in a parameter text refused (a final `Z` accepted), FlowLoadMode / PredictabilityMode checked (error 711, `J2D-CFG-007`). Known and left: a start
  entered in the same second as the last backlog entry gives no StartUtc (milliseconds dropped, the start must be strictly after): counted by a health field,
  the rule is AL's to decide; Propose rounds a sprint that starts at local midnight one day early (D7 of the audit).
- **Support codes.** Every error shown or logged carries `J2D-<AREA>-<NNN>` (`CFG`, `JIRA`, `NET`, `DATA`, `FLOW`, `ROVO`, `SYS`, `VBA`, `GEN`):
  the table is `src/jira2dashboard/error_codes.py` (code, vba number, meaning, cause, check, what support asks), shown in the sheet
  `ErrorCodes`; `modLog.ErrorCode(number, description)` is the VBA rule (602 has two causes: BaseUrl check, HTTP status = `JIRA-<status>`,
  status 0 = `NET-<WinHTTP error>`), `error_codes.code_of_vba` and `standalone.support_code` the Python one. A new `Err.Raise vbObjectError + n`
  needs its `Case n` in `ErrorCode` and a row in the table (`test_error_codes.py` fails otherwise); a code is never reused or renumbered.
  The scripts use `J2D-SCR-nnn` (every `Fail`, `throw` and unknown-argument message of `scripts\*.ps1` carries one: a test pins it), `diagnose` lists
  its `Findings` as `J2D-SYS-02x` with the line `Codes to quote to the support`, `check-machine` prints `J2D-SYS-040` on a FAIL.
- **Logs.** VBA (`modLog`): the version of the workbook on the `RUN` line (hidden name `J2D_Version`, written at build by
  `version.py`: commit and date, plus the version given by `release.ps1`), `PARAM` lines right after `RUN` (`LogParameters`: each Config row in clear
  but BaseUrl, Email, ApiToken, ProxyServer, ProxyBypass and Filter, which read `(set)`/`(empty)`; `SHOWN_PARAMETERS`, compared with
  `CONFIG_PARAMETERS` by `test_vba_log_guard.py`: a new parameter is masked until added there), the time at the end of each `END` line and of
  `RUN end`, a `RETRY` line before each new attempt of a call to Jira (`modHttp.RetryText`), the code on `RUN end stopped <code>`, and a `WARN J2D-DATA-003` line when status changes are left out because their date cannot be read (the count and the
  shape of the values, digits and letters masked: `modFlow.DateShape`, `mapping.date_shape`; an Empty date once made a negative DATE of a duration, which Excel
  refuses to write: error 7 "Out of memory" at the write of `StatusIntervals`). Python: the same,
  every Config value but the token. Keep a new step in a `LogBegin`/`LogEnd` pair; the Excel scenarios check the shape of a run (`logged_run` in
  `tests/steps/test_flow_excel.py`). `diagnose` quotes the end of the VBA log without its `PARAM` lines (it goes into a support ticket).
  Every Excel test (marker `excel`, all modules) leaves `out\inter\logs\excel-test_<test>.log` (fixture `keep_excel_logs` of `tests/conftest.py`,
  `testbench/keep_log.py`): result and time, its lines of each log of `out\inter\logs` (the VBA log, `vba-tests.log` of the VBA test runner), its
  `python.log`. `clean` keeps `out\inter\logs`, so a build does not delete them. `run.ps1` writes `scripts\logs\run_<command>.log` (one per command:
  the fake Jira of `start-mock` runs `run.ps1 mock-jira` in its own process) and sends the output of Python through PowerShell (`2>&1 | Show-Line`,
  UTF-8 console meanwhile), or the transcript would hold no result of pytest.
- **Propose values from Jira (button, VBA, AL 2026-11-13).** `modPropose` (`BtnProposeValues`, cell `A9`) runs `ReadJiraConfiguration` and then copies what fits
  of the proposal sheets (`_from_Jira`) into the parameter sheets (`_from_user`), **only into empty cells**: **since 2026-11-18 (AL) it no longer copies statuses or moves into
  `Workflows` / `Transitions`** (the metrics read them from Jira, `flow/status_rules.py`; those sheets hold exceptions), the `...Field` ids of `Config` that Jira names once (by the custom type, else by the name; two candidates = nothing;
  also `WorkTypeField` by the names "Work type category" / "TypeCategorie"). **Since 2026-11-18 (AL) the sheets `Workflows` and `Transitions` have no `Order` / `Type`
  column and the sheet `FieldMapping` is gone (nothing read them); a Team project with an empty `Team` is its own team (the project key: `ItemFlow.Team`,
  `modFlowBuild.TeamNames`, `flow_adapter`, `mapping.flow_issue`).** The Config lists of issue types and `BlocksLinkType` follow another rule (`SetListValue`): the Value is
  written when EMPTY or when it still equals its **Example** (the delivered default = never chosen), a typed value never changes, Jira agreeing = nothing: `SubTaskTypes`
  (JiraIssueTypes.Subtask), `FeatureTypes` (types of the Program project at its highest `HierarchyLevel`), `ParentRequiredTypes` / `PointsRequiredTypes` (a rule, not a fact:
  only the names that no Team project has are taken out), `BlocksLinkType` (the ONE link type of `JiraLinkTypes` whose outward text is "blocks", read by `modJiraConfig.ReadLinkTypes`,
  GET `issueLinkType`, a refused call gives none). A change of those defaults in `params/config.md` must keep Value = Example.
  Every cell written is yellow (`RGB(255, 242, 204)`), the first of a row has the note "To confirm: proposed from Jira". `Config.IsUserProjectAdmin` (True by
  default, False = no call to `workflows/preview`, which needs the project administrator right: `JiraStatuses` is still read, no moves).
  It also reads the sprints of the scrum boards (`modJiraConfig.ReadSprints`, agile/1.0, sheet `JiraSprints`; a 400/401/403/404 skips the board) and proposes the
  iterations they make (`ProposeIterations`: sprints that start and end within a day = one iteration, End = last day + 1, cut at the start of the next iteration so
  that no day is in two iterations, label = the shared name else `Iteration <Start>`, only the ones `Iterations` lacks). A new proposal rule goes in `modPropose`
  with its case in `T_ProposeFromSheets` / `T_ProposeTypes`. Not yet: the teams; PI dates have no API.
- **After an extraction (modHealth, AL 2026-11-13).** The statuses of the three buttons append `HealthFields` (Extract: only the column types, the
  tables come with Compute; Compute and Draw: every field), each field only when there is something to see:
  `Done outside the iterations: N of M` (no metric counts a ticket done outside every iteration: the usual cause of empty charts), `Without a team`,
  `Statuses not in Workflows` (their names), `Cells in error: Table.Column x n` (a formula that fails on real data; the charts' own #N/A are not read).
  Since 2026-11-22 (AL: "flow metrics are still not computable") also, in this order: `No start status: Feature, Team` and `No Wait status` first, then
  `Started in the second of the backlog entry`, `AsOf before the data`, `VelocityCheck gap`; they run in a log step `health`. **The Python twin is `flow/health.py`**
  (same fields, order and words, `tests/unit/test_health.py` reads `modHealth.bas`): the Python path writes them in its sheet `Run` (row Health) and `python.log`.
  The cause they name most: no IsStart status (Jira's workflows not read, none in Workflows) leaves every done ticket without StartUtc: Flow Time, Flow Efficiency
  and Flow Load not computable while Velocity counts; the log says it at the extraction (`INFO status rules | Feature: start n, active n, wait n, end n | Team: ...`,
  `WARN J2D-FLOW-010` / `J2D-FLOW-011`), `IsUserProjectAdmin` False skips Jira's workflows at the extraction too, a refusal is `WARN J2D-JIRA-030`; `WhyNotInFlowTime`
  says "No start status in the rules" or "No start after the backlog"; Flow Load is not computable (DailyLoad empty, never a false 0) when the workflow of the team
  has no start status. Numbers and status names only, never a value of a ticket; it never raises. The synchronous `ExtractAndCompute` (tests) does not write it. The fake data holds tickets
  done before the first iteration of the scenarios (9 of 26 on 2026-11-13): a "clean run" assertion checks the other fields (`Column types` included) only.
- **Column types of the raw sheets (AL, 2026-11-15).** After v0.7.0 wrote an empty date column and nobody saw it, the first FILLED cell of each column written from Jira
  (Issues, StatusIntervals, IssueLinks, Changelog; not the `Jira...` proposals) is checked against the type its label promises (a `...Utc` column holds a datetime). The spec is data:
  `flow/raw_types.py` (`RAW_TYPES`: Table, Column, Type, Required, Source, Allowed; a test pins that every raw column has a row and every `...Utc` is a datetime), written at build
  into the sheet `RawTypesSpec` (`Required` is `yes`/`no` text: `CStr(True)` is "Vrai" on a French PC) and read by `modHealth.TypeProblems`, called by `HealthFields` (also by the Python path:
  `standalone/__main__.log_column_types`). **One support code per data type:** `J2D-DATA-004` datetime, `005` date (midnight), `006` id (digits as text), `007` key (`ABC-12`), `008` integer,
  `009` number, `010` boolean, `011` enum (closed list: sprint state, status category), `012` list (no empty item), `013` text. A problem never stops the run: a status field
  ` | Column types: n column(s), first Issues.CreatedUtc expected datetime found number` and, in the log, three lines per problem with the same code: a WARN (column, expected, found, Jira's source
  path, first cell), an INFO (rows, filled, empty, the kinds of the first 20 filled cells, the SHAPE of the first value, the cell format on the VBA side), an INFO hint (a fixed sentence per
  type and kind found: 13 digits = epoch milliseconds...). Never a value: shapes only (`Double:9999999999999`). A required column with no filled cell is a problem too (`found empty`); an
  optional empty column is silent but shows in the `INFO types <Table> | rows N | filled: ...` line (continued on a next line: **the VBA log cuts a line at 300 characters**, `LOG_MAX_TEXT`,
  so a long message is several lines). `INFO wire types issue|configured fields|changelog | ...` says the JSON type Jira really sent for each field read (`modJira.WireTypesLines`,
  `mapping.wire_lines`): the line that shows a v0.7.0 at once. `HealthFields` is called BEFORE `LogRunEnd` so that its step and lines belong to the run. The hint sentences and the
  `FitsType`/`fits` rules exist on both sides: `test_raw_types.py` compares every `HintFor` literal with the Python hints. Not covered: only the first filled value is read (a bad value later
  in a column is not seen; the kinds of the first 20 are counted in the log). **The sheet `Changelog` is a plain range, not an Excel table** (`WriteTable`; the flow tables are
  made by `modFlowBuild.CreateTable`): `modHealth.TableShape` reads a table or, without one, the range from A1 (a first version read tables only and never saw `Changelog`,
  the column of the v0.7.0 failure: found by the two reviewer agents, hidden by a test that built a table the real workbook does not have: **a test of a reader builds the
  sheet as the real build does**). An empty table keeps one blank row (not a row). The filled cells are counted by `CountA` and the first 20 read by blocks of 2000 rows
  (`T_TypeBlocks`), so that 1 000 000 changelog rows cost almost nothing. The log lines come per table: its problems, then its fill lines (`raw_types.log_lines`).
- **Jira configuration (VBA).** `modJiraConfig` (button Read the Jira configuration, `BtnReadJiraConfiguration`) reads
  projects, issue types, statuses, workflows (`workflows/preview`) and fields into proposal sheets `Jira*`; it never writes
  the parameter sheets (calls and mapping: the header of `modJiraConfig.bas`). Besides the five first sheets it writes
  `JiraProjectWorkflows`, `JiraProjectFields` (createmeta, needs Create issues) and `JiraFieldUsage` (one page of issues per project). Its Rovo counterpart, to compare Jira with the workbook:
  `rovo/rovo_settings_analyser.md` (agent config, not code).
- **Prompt for Rovo (button, VBA).** `src/jira2dashboard/rovo_prompt.py` (pure text) is written at build into the sheet
  `RovoPromptSpec` (Part | Name | Text); `modParams.WriteRovoPrompt` (button `modActions.BtnRovoPrompt`, cell `A3`) turns it into the
  sheet `RovoPrompt`: Config read live (name, mandatory, meaning, example, never the Value), each sheet with its header, and the CSV
  format asked of Rovo (Sheet, Row, Column, Value, Note). No clipboard call (the VBA creates only two COM objects): the lines are selected.
  Rovo knows no data-check code: the Config descriptions (`params/config.md`, copied into the prompt as they are) name a check by its
  exact title, never DC01 (`tests/unit/test_rovo_prompt.py` pins it). **Since 2026-11-23 (AL) the prompt asks only what the workbook cannot read
  in Jira:** no connection parameter at all, not even named (spec part `keep`: `PROTECTED`, which the CSV load refuses; part `skip`: `CONNECTION_SETTINGS`
  IsUserProjectAdmin, TimeoutSeconds, MaxRetries and `FROM_JIRA`, the field ids, issue-type lists and BlocksLinkType that Propose reads), and no sheet
  Iterations (sprints) nor Transitions (moves): the sheets asked are Projects, PIs and the exceptions of Workflows (Workflow = Team or Feature). Part
  `load` (`LOADED_TOO`) keeps Iterations and Transitions loadable from a CSV (the configuration CSV of the fake Jira, `mock_workbook.CSV_SHEETS`).
  The button Load settings CSV (`BtnLoadRovoCsv`, cell `A5`: a dialog at each click, opened in the folder of the workbook, the CSV in any folder; `ImportRovoCsvFile(path)` is its work without the dialogs) reads Rovo's CSV
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
  COM: `flow/sheets.py`, `cli/check_machine.py`, `cli/diagnose.py`, `mockjira/__main__.py`, and the `build*`/`main` functions (see
  `pyproject.toml`).

Domain decisions and the metric definitions are in `docs/` (7 md files, **never edit them** unless AL says so: he allowed it on 2026-11-18 for the removal of `FieldMapping`, `Order` and `Type`). The handoff file `claude-code-handoff.md` was deleted
on 2026-11-06; its last version is `git show 6a93cdd:claude-code-handoff.md` (section 7: COM traps, 8b: open defect, 8c: Jira calls, and the roadmap).
`README.md` (English) is the engineering guide: keep it concise,
tables rather than code blocks for lists, never the word "client" nor the vendor or product name of the endpoint protection
(public repo: say EPDR or "endpoint protection"), prose lines <= 120 characters, ASCII for the architecture,
Mermaid for the workflows. Check a Mermaid edit with the real parser (`npm i mermaid jsdom` in the scratchpad,
then `mermaid.parse`), and check every `](#anchor)` link against the headings.

## Adding a derived table or a chart

1. Python: the formulas (`flow/formulas.py`: `*_COLUMNS` + `TABLES`), the reference (`flow/<name>.py`) with a hand-typed oracle in `tests/art_oracle.py`,
   the chart spec (`dashboard_spec.CHARTS`, before the daily curve), `blueprint.formula_rows` / `_chart_row`.
2. A new sheet name in the THREE lists (`sheet_names.EXCEL`, `modText.EXCEL_SHEETS`, `flow/sheets.error_cells`); with its suffix it fits the 31 characters of Excel
   (`_from_excel` takes 11: `test_sheet_names` fails otherwise, and the Python path stops at the write of its xlsx), then `modHealth` (cells in error),
   `DeleteDerived` and `BuildAll` of `modFlowBuild`.
3. Python path: `standalone/results.py` (+ its `Sheet(...)` in `__main__`) and `charts_svg.chart_pictures` (a new SERIES branch goes BEFORE the PI-axis one).
4. Pinned numbers: `test_run_vba_tests_pure` (markers, `FLB` range), `T_FlowBuild` ("N charts"), `test_blueprint` (`[CHARTS_SHEET] * n`), `test_dashboard_spec`
   (the chart list), `chart_blocks` / `_series` / the "== N" check of `tests/steps/test_flow_excel.py`, "N charts" in README, `manual.py`,
   `rovo/confluence_flow_metrics_page.md`, `tests/art_report.py` (the figure count).
5. Then `build` before `manual-test -Run`, `vba-reviewer` before `vba-tests`, and the generic `flow_excel.feature` compares the new table with the
   reference cell by cell (check that its data holds real figures, not only gaps).
6. **A column of `ItemFlow` reads only columns defined BEFORE it** (`test_flow_formulas` enforces it: `Team` read `[@IsFeature]`, defined later; use `check_formulas.in_list`
   inline). A formula typed in ONE cell of a table column fills the whole column: write per-row formulas (the `Reconciliation` Expected/Found) BEFORE `CreateTable`.

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
  Trivial entry points (`modActions`, `modButtons`) carry none. A new `.bas` module or `Cov` marker changes numbers pinned in tests:
  `test_run_vba_tests_pure` (modules, marker total, ranges of LOG/JCF/PRM/MEX/FLB), `test_child_env` (count of `*.bas`),
  `test_delivered_vba_com_objects` (`len(MODULES)`), and the lists `build_workbook.MODULES` and `run_vba_tests.ORDER`.
- **VBA conventions.** ASCII only, LF line endings, `Option Explicit`; offline tests `T_*` and fake-Jira tests `M_*` in
  `modTests.bas`; button macros are Public Subs without arguments. A `T_` test that writes dates into `TestScratch` ends with
  `ws.Cells.Clear`: `ClearContents` keeps the number format, and the next test reads 0/1 as dates. In a `T_` test never compare `CStr(Boolean)` or
  `CStr(Double)` ("Vrai", "1,5" on the French PC): use `IIf(x, "yes", "no")` or an Integer.
- **VBA traps.** A one-cell `Range.Value` is not an array; assigning an object to a Variant needs `Set` (error 450);
  `Application.OnTime` has a 1 s granularity; a COM call to an Excel that runs a macro is rejected (see `retry()` in
  `tests/steps/test_flow_excel.py`). `Empty - Date` is a negative DATE in VBA, and a `Range.Value = block` that holds one fails with
  error 7 "Out of memory" (it is no memory error): subtract through `CDbl` (`modFlow.DaysBetween`). A `Const` (or any module-level declaration) written
  after a procedure is COMPILE_FAILED (`--bisect` then blames an early module): declarations go at the top. A test cell that held a date keeps the date format:
  a number written there later (20261120) makes `.Value` fail before the code under test runs; set `NumberFormat = "General"` first.
- **VBA speed traps** (measured on the dev machine, 2026-11-07: the full extraction of the fake Jira, 1 244 issues, went from
  245 s to 24 s). `TypeName(object)` costs about 0.3 ms: use `TypeOf x Is Collection` or an `On Error` probe, never in a loop
  (`JGet` was 300 us, now 2 us). The first call on a fresh `Scripting.Dictionary` costs about 0.65 ms, even `Exists`
  (a `Collection` costs nothing), and `d.CompareMode = 0` on a new one about 0.3 ms: the default is already 0. Each JSON
  object is a Dictionary, so `JsonParse` still takes about 15 s of the 24 (a lighter JSON object is the next lever). `COUNTIFS`
  on a whole column (`Sheet!$A:$A`) scans its 1 048 576 rows (10 s for 7 284 cells; `SUMIFS`/`MATCH` stop at the used
  range, `COUNTIFS` does not): cut it at the last row with `refs.bounded` and a control cell (`test_no_countifs_reads_a_whole_column`).
  A step of the extraction job must stay short (about 1 s): `CHANGELOG_CHUNK` (50 issues) and `ROW_BATCH` (100 issues; a test
  sets `gRowBatch`). A slice writes as few cells as it can: clearing the 40 status cells at each slice brought the crash of
  the open defect back (3 runs in 5 on 2026-11-11), so `ShowProgress` touches the status rows only at the end of a run. Profile with the log (`BEGIN`/`END` of each sub-step) and a scratch copy of the workbook that gets a
  temporary benchmark module through `VBProject.VBComponents.Add` (never commit it).
- **Test markers.** `excel` and `network` tests are excluded from the default run; Excel scenarios build the delivered
  workbook once per module and open it in a fresh Excel per scenario. Each leaves `out\inter\logs\excel-test_<test>.log`.
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
- **Excel functions: only those of `tests/unit/test_blueprint.ALLOWED_FUNCTIONS`** (2026-11-24). The newest, TEXTSPLIT (lists of FeatureTypes since 2026-11-04),
  needs Microsoft 365 / Excel 2024: the client's Excel has it (its Velocity counts through it). Every test runs on a current Microsoft 365, so a function the
  client lacks only shows there as an empty list or a control to REVIEW (the IFERROR hides the #NAME?): add one to the list on purpose, after checking which Excel has it.
- Excel formula traps: a formula typed in one cell of a table column fills the whole column (write key columns as values
  before creating the table); a structured reference to a column whose name contains `_` is refused; compare dates with a
  half-second margin; a blank result is `""`, never `0`. Load settings CSV writes dates as ISO text (`SafeCell`): a formula
  that reads a date of the sheet PIs takes text too (`DATEVALUE`, as `formulas._pi_date` and DC18's `_date`).
- Mind the COM traps (section 7 of the deleted handoff, see above): positional arguments only, `put_en_us` for number formats,
  leading apostrophe on text written to cells (`as_cell` / `SafeCell`), read `.Value2`, `Comment.Text` is a call in pywin32 (`.Text()`), activate a chart's sheet and its ChartObject before export (else an empty PNG).
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
  Since 2026-11-05 the event log names `ntdll.dll`, code `0xc0000028` or `0xc00000ff` (a failed stack unwind), offset
  `0x1d3df`; it also hits `BEGIN flow sheets: charts` of a second build in the same Excel (scenario "second calcul": 1 in 5
  at 61507d1 on 2026-11-12, `T_FlowBuild` alone 0 in 10). Measured as this defect then; still run the A/B above before
  dismissing a new crash there. `vba-tests` loses Excel at `BEGIN flow sheets: charts` (T_FlowBuild) too: 4 runs in about 14 on
  2026-11-13, none tied to a change: rerun up to 3 times.
- **Jira API documentation:** the Atlassian MCP is not connected (only its authenticate tools) and a real site is never authenticated (hard constraint). Read the
  official OpenAPI schema with Context7 (`/openapi/dac-static_atlassian_cloud_jira_platform_swagger-v3_v3_json`, and Jira Software `.../swagger.v3.json`) or WebFetch of
  `https://dac-static.atlassian.com/cloud/jira/platform/swagger-v3.v3.json`: the developer.atlassian.com pages need JavaScript.
- A leftover Excel of a crashed run (title "Excel", started during the run) is stopped with PowerShell `Stop-Process -Id <pid>`: `taskkill /PID` is mangled by Git Bash.
- **Long runs in the background:** `cmd &` inside a background call notifies at once, not at the end: wait with
  `until grep -qE "passed|failed" log; do sleep 15; done` in a background Bash. The harness kills background shells when the machine is low
  on memory: run the test groups one after the other, and rerun a killed one only when asked.
- **Git Bash rewrites `/switch` arguments** of Windows programs into paths (`/update` -> `C:/.../Git/update`, also
  `wevtutil /q:`): prefix the command with `MSYS_NO_PATHCONV=1`.
- The PowerShell tool is slow here (CIM / event-log queries time out): use Bash `tasklist` for processes.
- Reading PowerShell output from Python (`Get-Help`, the scripts): `encoding="oem", errors="replace"`, not UTF-8.
- **A Bash call with a heredoc that holds apostrophes (French text, VBA comments, `\'`) can fail with "unexpected EOF while looking for
  matching `'`" and apply NOTHING** (2026-11-16, several times): write the edit script with the Write tool into the scratchpad, then run it
  (`python3 <path>`); a heredoc whose body is plain code with no apostrophe is fine. Also never name a test module like another one of
  `tests/` (`tests/steps/test_art_metrics.py` vs `tests/unit/test_art_metrics.py` broke the collection: "imported module has this
  `__file__`"): the Python-path steps are `test_art_metrics_python.py`.
- **A VBA error in a `T_` test or in the build of a test workbook opens the dialog "Microsoft Visual Basic" and hangs `vba-tests`** (no
  `summary:` line; the log ends at a `BEGIN`): stop that Excel by its PID (title + start time of the run), read the end of
  `out\inter\logs\vba-tests.log`. The cause on 2026-11-16: `PrepareRaw` of the new raw sheet `PIIntervals` found it empty because the
  `FlowRaw` fixture of `T_FlowBuild` did not write it: a new raw table needs its fixture written by hand in `modTests`.
- A local variable must not have the name of a procedure of the same module (`piRows` hid `PiRows`: VBA ignores case); a check of a test that
  reads Excel cells goes through a function that returns "?" on an error or an error value (`IterationCell`, as `AxisEnds`).
- Edit scripts with quotes, `\r\n` or backslashes (Windows paths in Markdown): write a `.py` file with the Write tool and
  run it, not a Bash heredoc (a heredoc turned `tests\\features` into a form feed in a Markdown file).
  Pattern: `assert s.count(old) == 1` before each replace, write with `newline="\n"`, run with `.venv/Scripts/python.exe`.
- In those `.py` edit scripts, VBA text with `""` ends a `"""` string early and `"\vba-tests.log"` holds a `\v`:
  use `r'''...'''` and run `python -m py_compile` on the script first. `\r`, `\b`, `\t`, `\v` in `out\release`, `\bootstrap`...
  become control characters without a warning: raw strings, then `grep -cP '[\x00-\x08\x0b-\x1f]' file` must give 0. Create the
  script with the Write tool, not a Bash heredoc (one ended at end-of-file and applied nothing).
- While `test-excel` / `vba-tests` run in the background, do not edit `src/` or `vba/`: each Excel module builds its
  workbook from `vba/*.bas` when it starts, and the child processes import `src`.
- Pester fakes of python are functions: a line they write to stderr lands in the transcript as a decorated error record,
  which `python.exe` called with `2>&1 | Show-Line` does not do. Assert the text, not the absence of decoration.
- `Windows fatal exception: code 0x800706be` (or `0x800706ba`) printed after Excel scenarios is teardown noise, also in
  green runs: only the test result counts.
- **While a change adds `Cov` markers,** the hook `bas_guard` prints "VBA guard tests failed" after each `.bas` edit until
  `test_run_vba_tests_pure` pins the new marker total and `FLB`/`MEX`... range: expected; finish the edits, then update those two numbers.
- **Removing or adding a column of a parameter sheet** (`params/*.md`) shifts two things silently: the tuples read by INDEX in Python (`row[2]` in `mockjira/`, `flow/model.py`, `checks.py`, tests) and
  the formulas that read the sheet by COLUMN LETTER (`formulas.INTERVAL_HELPERS` on `Workflows_from_user!$D:$F`): only the `flow_excel.feature` scenarios catch the second (empty `StartUtc`). grep both.
- **A `T_` check on a range** must not raise on an error cell: `Application.WorksheetFunction.Sum` of a column with `#N/A` raises (opens the VB dialog): use `ColumnSum` (returns "?").
  A check whose only possible answers are "1" or "?" proves nothing (an `IFERROR(...,"")` formula hides its failure): seed a ticket or leave it to the Excel scenarios.
- **Look at the results of a scenario without a click:** `PYTHONPATH=src .venv/Scripts/python.exe -m jira2dashboard.mock_workbook --scenario art-metrics --run` (needs `build` first), then
  `standalone.xlsx_read.read_sheets(Path(copy))` shows the cached values of every sheet. A `python -` heredoc that imports `jira2dashboard` needs `PYTHONPATH=src`.
- **AL's Excel is French and shows dates as `45474,49546`:** to give him a filter, convert with `(date - dt.date(1899, 12, 30)).days` (2026-10-28 = 46323) and say that the End is excluded.
- **Mutation proof of a Python reference:** `cp file $TEMP/x.bak`, replace one expression (`assert s.count(old) == 1`), run the unit test file, `cp` back; one mutation per behaviour. A mutation that
  survives means the test data cannot tell it (a tautology: add the case, or let the function take the ledger as a parameter so a test can feed a wrong one).
- **A scenario must not hold what Jira gives in its parameter sheets** (2026-11-18): `Dataset.jira_sheets` = Jira's full Workflows/Transitions, `sheets` = the exceptions. References
  come from `flow_adapter.jira_rules(ds)` / `jira_workflows(ds)`, never `Rules.from_params(list(ds.sheets))`; a test whose dataset still has complete sheets cannot see an exceptions-only workbook.
- **A sheet that a formula reads must exist from the build** (`flow/sheets.py`, list of `add_sheet_after_last`), even empty: the formula written at the build of the flow sheets would
  otherwise ask Excel for a file. A new ItemFlow column goes LAST (`formulas.ITEMFLOW_COLUMNS` after `ITEMFLOW_CHECK_COLUMNS`, `results.ITEMFLOW` and the positional tuple of `itemflow_rows`).
- **A new `Cov` marker takes the next free number of its prefix** (`grep -o 'Cov "FLB[0-9]*"' vba/*.bas | sort`): FLB50-57 were taken when new ones were numbered FLB53-55 by mistake. Pins: total +
  range of the prefix in `test_run_vba_tests_pure.py`. `StatusNames(wb)` and any build-time reader see the RAW sheets only (formulas are not calculated yet).
- The scenario "Annuler pendant l'extraction" is the usual victim of the Excel crash (open defect): rerun it alone with `-k annuler` (30 s) before suspecting a change; green twice = noise.
- `jira_config.feature` Propose scenarios assert that a deleted Workflows/Transitions row does NOT come back (Propose no longer copies them): do not "fix" that back.
- **Where the time of an Excel test goes:** the first line of `out\inter\logs\excel-test_<test>.log` is `# test | result | time`; `run.ps1 test-excel ... --durations=0` adds setup and teardown
  (that is how the 5.3 s teardown showed). cProfile tells nothing here (the main thread waits in COM; `select` is the fake Jira's thread): time the steps with `date +%s` or the VBA log `END ... | n s` lines.
- **Excel never leaves while the caller holds `app`** (it always does inside `with excel_session()`), and needs about 3 s to leave even without a reference: `com.EXIT_GRACE_S = 0.5` ends
  this session's own process (pid + `excel.exe` checked) right after Quit; waiting 5 s only delayed the same kill, 5 s for each of the 66 Excel tests. `check-machine` keeps 5 s.
- **Timings of the full pass (2026-11-19, 14 cores):** coverage 73 s, vba-tests 41 s, test-excel flow `-n 4` 119 s, the other Excel modules 77 s, test-scripts 104 s, build 9 s = about 7 min;
  `test-excel` flow without `-n` is 290 s. A step far above that is worth a look (a leftover Excel, a machine under load: the first symptom is every test 2 to 3 times slower).
- A line of a `.ps1` or a doc that holds a Windows path (`scripts\run.ps1`) goes through the Edit tool, not a `python -` script: `\s`, `\i` are escapes there (SyntaxWarning, and the `assert s.count(old) == 1` then fails).
- **Look at a chart:** `start-mock.ps1 -Port 8100`, a scratch copy of `out\release\jira2dashboard.xlsm` filled by
  `mock_workbook.fill_copy(copy, base_url(8100))`, `modActions.ExtractAndCompute`, then in the same session Activate the
  ChartObject, `Chart.Export` to the scratchpad and Read the PNG; `stop-mock.ps1 -Port 8100` after (8099 may be AL's).
- **Look at the charts page** (`out\manual_test\charts\<scenario>\index.html`): the Playwright MCP refuses `file:`. Serve the folder with
  `( cd out/manual_test/charts && .venv/Scripts/python.exe -m http.server 8123 --bind 127.0.0.1 & )`, open `http://127.0.0.1:8123/<scenario>/index.html`,
  `browser_take_screenshot` with a `filename` under `C:\dev\jira2dashboard\.playwright-mcp\` (other roots are refused), Read it; then stop the
  server by its PID, `browser_close`, `rm -rf .playwright-mcp` (the tool leaves that folder in the working directory).
- **A Bash call that `cd`s away from the repo moves the working directory of the session** (a `cd` into `out\manual_test\charts` put the
  browser tool's `.playwright-mcp` folder there): start each Bash command with `cd /c/dev/jira2dashboard`, or use a subshell `( cd x && ... )`.
- **Hung test Excel (VBA error dialog):** run `vba-tests` with `timeout 420`. The Excel to kill is the one whose window
  title is "Microsoft Visual Basic" and whose start time matches the run (`Get-Process EXCEL | Select Id, StartTime,
  MainWindowTitle`), never an older one. Mutation proof of a VBA test: copy the module to `$TEMP`, mutate it with a Python
  script (`assert count == 1`), run, restore with `cp` in the same command. A new check that reads Excel objects must not
  raise (a raise opens that dialog): wrap it like `AxisEnds` in `modTests.bas`.
- **Excel's writes undone (a machine with an endpoint protection):** `build` prints `built ...` then Excel cannot open the file, or
  `mock-workbook` leaves the copy empty (`J2D-SYS-012` / `J2D-SYS-013`, then `J2D-CFG-002` in the scenario), while a workbook saved by
  hand stays and the macros of a downloaded release run. `run.ps1 check-machine` tests it (`save: ...`, 3 s after the write). Fix: an
  allow-list on the repository folder for `python.exe` and `EXCEL.EXE`; or use the workbook of a release, no build needed.
- **Excel hygiene (AL's rule):** after every Excel run, even a green one, `tasklist | grep -c EXCEL` must be back to what it was;
  stop only the Excel your run started (start time; title "Microsoft Visual Basic" when hung), never an older one.
- The Excel that `manual-test.ps1` opens is AL's: leave it open; when AL says done, check `tasklist`, then `stop-mock.ps1 -Port 8099`.
- `com_error -2146959355` ("Echec de l'execution du serveur") at `DispatchEx` = Excel failed to start: transient, rerun.
- Excel rounds the position of a form control down to a whole point: a column of 124.5 pt put every button one cell to the
  left (`TopLeftCell` N1 for O1). `com.add_macro_button` rounds up.
- After an Excel crash in a test, Office starts a new `EXCEL.EXE` (title "Excel" or "... - Récupération") at the same second:
  it belongs to the run; stop it by its PID once its start time is checked.
- **A few VBA tests only:** copy `modTests.bas`, cut `TestNames()` to the tests wanted, run `vba-tests --offline`, restore with `cp`
  (`vba-tests` takes about 1 minute whole, measured 2026-11-18: rarely worth it; never commit the cut list).
- **Mutation of a VBA test:** no `summary:` line in the log = no proof (a mutation that raises inside a test, e.g. a type mismatch, loses the run:
  pick one that changes a result without raising). `Check` takes (area, label, input, expected, actual); the log cuts a line at 300 characters.
- pywin32: `Range.Address` is a property (`cell.Address.replace("$", "")`), not a call with arguments; `ws.Cells(1, n).Value` reads a one-cell range as a scalar.
- **`test-excel` by groups** (the whole run was killed once for memory, leaving an orphan `EXCEL.EXE`: stop it by PID): `tests/steps/test_build_workbook.py` (2 min),
  `test_macro_check.py` + `test_manual_test.py` (2 min), `test_params.py` + `test_standalone.py` (2 min), `test_flow_excel.py` (2 min with `-n 4`, 5 min without; foreground, `timeout 595`).
  **`-n 4` (pytest-xdist) is for the flow module only** (2026-11-19, 5 runs in 5 green, 130 s against 290 s): each worker builds and opens its own workbook `out\interlow_gwN\` (`WORKER` in `test_flow_excel.py`),
  so the VBA log and the files of a scenario are never shared; the other modules share workbook paths, never give them `-n`. `com.start_excel` tries again on "Server execution failed" (several Excel starting together).
  `-k` matches function AND module names: `-k scenario` ran all 26 tests of `test_flow_excel.py` (5 min); the two chart scenarios are
  `-k "art_metrics or png or PNG or graphiques"` (30 s). Read the `N passed` count: a `-k` that selects nothing or too much looks green.
- **Calling a `scripts\*.ps1` with optional arguments: do not splat.** A hashtable splat shows as `-Port: 8100` in the Pester fakes
  (`Calls`), and an array of `"-Port", $p` strings is caught by `$Rest` (J2D-SCR-001). Write the two calls:
  `if ($Scenario) { & x -Port $Port -Scenario $Scenario } else { & x -Port $Port }`. A default that depends on another parameter is
  `[int]$Port = 0` then `if ($Port -eq 0) { $Port = ... }` (`$ScenarioPort` = `Scenario.port`, pinned by a test).
- **Read the parameter sheets of a saved copy without Excel:** `standalone.params_book.read_params(Path(copy))` (a zip read; the keys are the
  LOGICAL names `Config`, `Projects`, `Iterations`, `PIs`, `Workflows`, `Transitions`, not `PIs_from_user`); `xlsx_read.read_sheets` reads any
  sheet by its real name. Use it to check what a run left in `out\manual_test\<scenario>\` (the calendar was checked so on 2026-11-17).
- In an edit script a backslash goes through `chr(92)` or a raw string; after the run `python -m py_compile` the target and
  `grep -cP '[\x00-\x08\x0b-\x1f]'` it (a `\b` in a plain triple-quoted string became a backspace, a `\r\n` a real line break).
  That grep misses the escape still written in the SOURCE of a Python file: `scripts\run.ps1` typed in a `print` string compiles without a
  warning (`\r` is valid) and prints a carriage return; only `\m` is flagged (`python -W error -m py_compile file`). Also run
  `grep -nP '[A-Za-z]\\[rntbvf][A-Za-z]' file`; write such a path with `/`, `chr(92)`, a raw string, or with the Edit tool.
- `_launch.cmd` traps: a `for` list drops `/?` (wildcard), so it is tested apart; a launcher starts the helper without `call`
  (`call` swallows `/?`); from a test, `cmd /c` needs `.\name.cmd`.
- After moving code between `.bas` modules: `Private` names (Const, Function) used from the new module break the compile of
  the test workbook only (`build` compiles the delivered modules, `vba-tests` adds `modTests`). `--bisect` loads the
  modules in order, so a forward reference shows as a false failure at an early module: scan for `Private` names instead.
  Python steps call `modX.Name` explicitly (`tests/steps/test_params.py`): update them.
- The VBA guard hook also runs after a Bash/PowerShell command that changed a `.bas`, but only from the second command
  of a machine (the first one records the modules).
- `test_vba_log_guard.py` reads the whole line of a log call, its trailing comment too: "URL", "body" or "resp" in a
  comment on a `LogLine` line fails it.
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
A manual run on 2026-11-11 (log with the time of each step) took 16.8 s for the same 1 244 issues: 13 search pages
about 0.3 s each, 25 changelog chunks about 0.3 s each, ItemFlow formulas 1.9 s. Re-measure the share of `JsonParse`
before starting (the 2026-11-07 figures above predate it).
Also pending, offered to AL: when the macro loses its state (Esc then End during a long step), the Dashboard stays on its
last status and Cancel does nothing; `AbortExtraction` and a slice that finds `mBusy` False could write a log line and
"Stopped" in `A16` (`dashStatus`).
