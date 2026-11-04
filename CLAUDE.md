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
.\scripts\run.ps1 mock-jira       # fake Jira on http://127.0.0.1:8099
.\scripts\run.ps1 build           # delivery workbook out\jira2dashboard.xlsm (--out name.xlsm)
```

`test` and `coverage` skip the `excel` and `network` markers by default (`addopts` in `pyproject.toml`): use `test-excel` /
`test-network` for those.

Single test: `.\scripts\run.ps1 test tests/unit/test_checks.py::test_name` or
`.\scripts\run.ps1 test-excel tests/steps/test_flow_excel.py -k "contr" -q`. Call `run.ps1` directly from PowerShell: do not
wrap it in `cmd /c` or redirect with `2>&1` (it stops on `$ErrorActionPreference='Stop'`).

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
- **Extraction (VBA).** `modHttp` (WinHTTP, proxy with Windows auth, retries) -> `modJira` (search, bulk changelog) ->
  `modExtract` -> `modFlow` (status intervals) ; `modActions` is the dashboard button (`ExtractAndCompute`). The previous
  results are cleared only after Jira has answered.
- **Fake Jira.** `mockjira/` is a stdlib `http.server` Jira Cloud (search/jql, changelog bulkfetch, proxy/CONNECT, fault
  injection) fed by `dataset.py` (hand-written cases via `issue()`) and `generator.py` (seeded volume);
  `anomalies.py` adds known data-quality cases. The real Jira is never called.
- **COM layer.** `excel/com.py` holds the pure helpers and `excel_session`; `testbench/` builds, compiles and runs the VBA
  tests; `testbench/children/` scripts run as separate processes by path. Files omitted from coverage because they need
  COM: `flow/sheets.py`, `cli/check_machine.py`, `mockjira/__main__.py`, and the `build*`/`main` functions (see
  `pyproject.toml`).

Domain decisions and the metric definitions are in `docs/` (7 md files, **never edit them**: their MD5 is tracked in
`claude-code-handoff.md`, which also holds the roadmap, review dispositions and the MD5 of every `.bas`). Update the MD5
table in the handoff whenever a `.bas` changes. `README.md` (French) covers delivery to the client machine.

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
- **Test markers.** `excel` and `network` tests are excluded from the default run; Excel scenarios share one delivered
  workbook per module.

## Hard constraints

- **Nothing from the client may leave the client machine** (no scripts, no data). Client values (`BaseUrl`, `Email`,
  `ApiToken`, `ProxyServer`, `ProxyBypass`, `Filter`, `WorkTypeField`, `AsOf`, the periods calendar) stay **empty** in the
  repo; the API token must never be written to `params/config.md` (a test refuses it).
- **Never kill an `EXCEL.EXE` the scripts did not start** (the user may have their own). If a test run hangs, end only the
  process you started, by PID. A modal VB dialog is the usual cause of a hung Excel run.
- **Do not call a real Jira.** Use the fake Jira.
- Excel formula traps: a formula typed in one cell of a table column fills the whole column (write key columns as values
  before creating the table); a structured reference to a column whose name contains `_` is refused; compare dates with a
  half-second margin; a blank result is `""`, never `0`.
- Mind the COM traps listed in `claude-code-handoff.md` section 7: positional arguments only, `put_en_us` for number formats,
  leading apostrophe on text written to cells (`as_cell` / `SafeCell`), read `.Value2`, activate a chart's sheet before export.
- **Git.** Repo `AlexisAIchangeAgent/jira2dashboard` (public), branch `main`. `.gitattributes` disables line-ending
  conversion (`* -text`, `*.bas text eol=lf`) so the tracked MD5 hold; Excel binaries (`*.xlsm`, `out/`) are ignored.
  Commit messages end with `Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>`. If SSH push is refused, push over
  HTTPS with the `gh` token: `git -c credential.helper= -c "credential.helper=!gh auth git-credential" push https://github.com/AlexisAIchangeAgent/jira2dashboard.git main`.

## Roadmap not done yet (ask before starting)

Progress bar and cancel during extraction. Real test at the client machine is still to be done by AL.
