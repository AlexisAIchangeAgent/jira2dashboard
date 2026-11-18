"""Prepare a copy of a workbook to try it by hand against the fake Jira (scripts\\manual-test.ps1 runs it).

The copy, out\\manual_test\\jira2dashboard4manualtest.xlsm, gets what a client would type: in Config the address,
account, reference date and custom field ids of the fake Jira of `run.ps1 mock-jira`; 14-day iterations in Iterations;
the two PIs of the fake data in PIs. The workbook given is never changed: the one to deliver keeps its client values
empty. out\\manual_test is never emptied by clean, build or vba-tests; the copy is rewritten at each run.
With --scenario the copy is the one of that scenario alone, in out/manual_test/<scenario>/: its own Projects, Workflows,
Transitions, Iterations and PIs (nothing of params/*.md), its fake Jira on its own port; --run also extracts and exports its charts.
Next to it, rovo_mock_jira.csv: the same configuration as Rovo would answer it (the CSV of the button Load settings CSV), to
try that button by hand. The connection (address, account, token, proxy) is not in it: the button never loads it.
"""
import csv
import datetime as dt
import io
import shutil
import sys
import time
from pathlib import Path

from jira2dashboard import rovo_prompt as RP
from jira2dashboard import sheet_names as S
from jira2dashboard.cli.args import parser
from jira2dashboard.excel import com
from jira2dashboard.flow import dashboard_spec as D
from jira2dashboard.flow.model import Iteration
from jira2dashboard.mockjira.anomalies import PI, PI_WINDOW
from jira2dashboard.mockjira import scenarios
from jira2dashboard.mockjira.dataset import FieldIds, Scenario
from jira2dashboard.mockjira.generator import AS_OF, SINCE
from jira2dashboard.mockjira.generator import PIS as GENERATED_PIS
from jira2dashboard.mockjira.server import MOCK_EMAIL, MOCK_TOKEN, MockJira
from jira2dashboard.params.model import load_all
from jira2dashboard.paths import CHARTS_DIR, MANUAL_TEST_DIR, RELEASE_DIR, REPO_ROOT
from jira2dashboard.standalone.params_book import read_params
from jira2dashboard.standalone.xlsx_write import Sheet, write_xlsx

COPY_NAME = "jira2dashboard4manualtest.xlsm"
COPY = MANUAL_TEST_DIR / COPY_NAME
CSV_NAME = "rovo_mock_jira.csv"
CSV_NOTE = "configuration of the fake Jira (run.ps1 mock-jira)"
FIELD_IDS = {"StoryPointsField": "story_points", "BusinessValueField": "business_value", "PIField": "pi",
             "TeamField": "team", "SprintField": "sprint", "TargetStartField": "target_start",
             "TargetEndField": "target_end"}          # Config parameter of a data check -> attribute of FieldIds
# the PI of the anomalies, then the other PI the generator draws (AL, 2026-11-12: two bars on Feature Predictability). It
# draws either PI over the same sprints, so both get the window of the anomalies: DC18 judges them alike.
PIS = [Iteration(PI, *PI_WINDOW)] + [Iteration(name, *PI_WINDOW) for name in GENERATED_PIS if name != PI]
ROWS = "A2:C200"                                      # the rows of Iterations and PIs that the copy rewrites


def base_url(port: int) -> str:
    return f"http://127.0.0.1:{port}"


def config_values(url: str, fields: FieldIds = FieldIds(), as_of: dt.datetime = AS_OF) -> dict[str, str]:
    """Parameter -> value of the Config sheet, for the fake Jira at `url`."""
    return {"BaseUrl": url, "Email": MOCK_EMAIL, "ApiToken": MOCK_TOKEN, "AsOf": f"{as_of:%Y-%m-%d %H:%M}",
            "WorkTypeField": fields.work_type or ""} | {name: getattr(fields, attr) or "" for name, attr in FIELD_IDS.items()}


def copy_path(scenario: Scenario | None = None) -> Path:
    """The copy to open: the generic one, or the one of a scenario in its own folder out\\manual_test\\<scenario>\\ (a separate
    instance: its logs, its Rovo CSV and its results are not mixed with the generic copy)."""
    if scenario is None:
        return COPY
    return COPY.parent / scenario.name / f"jira2dashboard4{scenario.name}.xlsm"


def csv_path(scenario: Scenario | None = None) -> Path:
    return COPY.parent / CSV_NAME if scenario is None else copy_path(scenario).parent / f"rovo_{scenario.name}.csv"


def iterations(first: dt.datetime = SINCE, as_of: dt.datetime = AS_OF, days: int = 14) -> list[Iteration]:
    """I1, I2... of `days` days from `first`, one after the other, the last one holding `as_of`."""
    iterations, start = [], first
    while start < as_of:
        end = start + dt.timedelta(days=days)
        iterations.append(Iteration(f"I{len(iterations) + 1}", start, end))
        start = end
    return iterations


def scenario_config(scenario: Scenario | None, url: str) -> dict[str, str]:
    """config_values for the fake Jira at `url`: the generated data, or the fields and reference date of a scenario."""
    return config_values(url) if scenario is None else config_values(url, scenario.dataset.fields, scenario.dataset.as_of)


def params_book(url: str, values: dict[str, str] | None = None, scenario: Scenario | None = None) -> list[Sheet]:
    """The parameter sheets of params/*.md filled for the fake Jira at `url`, as sheets of an xlsx (no Excel): Config
    with config_values (and `values` over them), the iterations in Iterations and PIS in PIs, as
    dates. What the Python extraction reads from a filled workbook. A scenario brings its own PIs, iterations, Workflows
    and Transitions."""
    filled = scenario_config(scenario, url) | (values or {})
    out = []
    for sheet in (load_all() if scenario is None else scenario.dataset.sheets):
        rows = list(sheet.rows)
        if sheet.name == "Config":
            value = sheet.header.index("Value")
            rows = [row[:value] + (filled.get(row[0], row[value]),) + row[value + 1:] for row in rows]
        elif sheet.name in ("Iterations", "PIs"):
            dated = iterations() if sheet.name == "Iterations" else PIS
            if scenario is not None:
                dated = scenario.iterations if sheet.name == "Iterations" else scenario.pis
            rows = [(p.name, p.start, p.end) for p in dated]
        out.append(Sheet(S.of(sheet.name), sheet.header, rows))
    return out


def write_params_book(path: Path, url: str, values: dict[str, str] | None = None, scenario: Scenario | None = None) -> Path:
    return write_xlsx(path, params_book(url, values, scenario))


def csv_value(value) -> str:
    """A cell of a parameter sheet as Rovo writes it: a date as yyyy-mm-dd (with the time when it has one)."""
    if isinstance(value, dt.datetime):
        return f"{value:%Y-%m-%d}" if value.time() == dt.time() else f"{value:%Y-%m-%d %H:%M}"
    return "" if value is None else str(value)


def rovo_csv(port: int = 8099, scenario: Scenario | None = None) -> str:
    """The answer Rovo would give for the fake Jira, in the CSV the button Load settings CSV reads (rovo_prompt.CSV_COLUMNS):
    the Config values of config_values but the connection (rovo_prompt.PROTECTED, never loaded from a CSV), then every row
    of the sheets the prompt asks for (rovo_prompt.SHEET_RULES), with a BOM as Excel writes a UTF-8 CSV."""
    out = io.StringIO()
    writer = csv.writer(out, lineterminator="\r\n")
    writer.writerow(RP.CSV_COLUMNS)
    for name, value in scenario_config(scenario, base_url(port)).items():
        if name not in RP.PROTECTED:
            writer.writerow((S.of("Config"), "", name, value, CSV_NOTE))
    for sheet in params_book(base_url(port), scenario=scenario):
        if S.logical(sheet.name) in RP.SHEET_RULES:
            for number, row in enumerate(sheet.rows, start=1):
                for column, value in zip(sheet.header, row):
                    writer.writerow((sheet.name, number, column, csv_value(value), ""))
    return "\ufeff" + out.getvalue()


def write_iterations(ws, iterations: list[Iteration]) -> None:
    """Rows Name | Start | End of the sheet Iterations or PIs; the rows of an earlier run are cleared first."""
    ws.Range(ROWS).ClearContents()
    for row, iteration in enumerate(iterations, start=2):
        ws.Cells(row, 1).Value = com.as_cell(iteration.name)
        ws.Cells(row, 2).Value2 = com.to_excel_serial(iteration.start)
        ws.Cells(row, 3).Value2 = com.to_excel_serial(iteration.end)


def write_config(ws, values: dict[str, str]) -> None:
    """Column Value of the Config rows named in column Parameter."""
    rows = {ws.Cells(r, 1).Value: r for r in range(2, ws.UsedRange.Rows.Count + 1)}
    for name, value in values.items():
        ws.Cells(rows[name], 2).Value = com.as_cell(value)


def write_rows(ws, rows) -> None:
    """The rows of a parameter sheet (Workflows, Transitions) under its header; the rows of an earlier run are cleared first."""
    ws.Range("A2:H200").ClearContents()
    com.write_block(ws, 2, 1, [list(row) for row in rows])


def fill_workbook(wb, url: str, scenario: Scenario | None = None) -> None:
    """What a client types, into the open workbook `wb`: Config, Iterations, PIs; for a scenario also its own Projects, Workflows and
    Transitions (the default rows of the delivered workbook are replaced: nothing of them stays). Used by the copy and by the tests."""
    write_config(wb.Worksheets(S.of("Config")), scenario_config(scenario, url))
    write_iterations(wb.Worksheets(S.of("Iterations")), iterations() if scenario is None else list(scenario.iterations))
    write_iterations(wb.Worksheets(S.of("PIs")), PIS if scenario is None else list(scenario.pis))
    if scenario is not None:
        for sheet in scenario.dataset.sheets:
            if sheet.name in ("Projects", "Workflows", "Transitions"):
                write_rows(wb.Worksheets(S.of(sheet.name)), sheet.rows)


def fill_copy(path: Path, url: str, scenario: Scenario | None = None) -> None:
    with com.excel_session() as app:
        wb = com.open_without_macros(app, path)
        fill_workbook(wb, url, scenario)
        wb.Save()
        wb.Close(False)
        del wb, app                                                   # release every COM proxy before Excel quits


SETTLE_S = 3        # seconds to wait before reading the copy back: a protection that undoes a write does it a moment later


def holds_the_fill(path: Path, url: str) -> bool:
    """Does the saved copy hold the address that fill_copy wrote? Read as a file (no Excel), as the extraction reads it."""
    config = next(sheet for sheet in read_params(path) if sheet.name == "Config")
    return {row[0]: row[1] for row in config.rows}.get("BaseUrl") == url


EXAMPLES = """
examples:
  scripts\\run.ps1 mock-workbook                                 the copy of out\\release\\jira2dashboard.xlsm
  scripts\\run.ps1 mock-workbook --workbook C:\\x.xlsm --port 8100  another workbook, a fake Jira on port 8100
"""


def parse(argv: list[str] | None = None):
    p = parser("scripts\\run.ps1 mock-workbook", __doc__, EXAMPLES)
    p.add_argument("--workbook", default=str(RELEASE_DIR / "jira2dashboard.xlsm"),
                   help="the workbook to copy (default out\\release\\jira2dashboard.xlsm); it is never changed")
    p.add_argument("--port", type=int, default=None,
                   help="port of the fake Jira on 127.0.0.1 (default 8099; a scenario has its own: 8101 for art-metrics)")
    p.add_argument("--scenario", choices=sorted(scenarios.SCENARIOS), default=None,
                   help="fill the copy for a scenario of the fake Jira (mock-jira --scenario): its calendar, projects, workflow and "
                        "fields, in its own copy out\\manual_test\\<scenario>\\jira2dashboard4<scenario>.xlsm")
    p.add_argument("--run", action="store_true",
                   help="with --scenario: also run it unattended (serve it, extract with the macros of the copy, export the charts "
                        "as SVG and PNG with an index.html, save the copy with its results, close Excel)")
    return p.parse_args(argv)


def status_of(wb) -> str:
    """What the run left in the workbook, for the log of the run: the rows of the table Issues and the Dashboard headline; "?" when
    it cannot be read (a COM failure here must not hide the run's own result)."""
    try:
        rows = wb.Worksheets(S.of("Issues")).ListObjects("Issues").ListRows.Count
        return f"{rows} issue(s) in the sheet {S.of('Issues')}, status {wb.Names(D.STATUS).RefersToRange.Value}"
    except Exception:
        return "?"


def report_builder():
    """The function that writes the report expected / obtained of a scenario (tests/art_report.py: it holds the numbers worked out by
    hand, which stay in tests/), or None when this is not a checkout with its tests (the report is then left out)."""
    import importlib.util
    import sys
    path = REPO_ROOT / "tests" / "art_report.py"
    if not path.is_file():
        return None
    sys.path.insert(0, str(path.parent))                         # art_report imports art_oracle, which sits beside it
    try:
        spec = importlib.util.spec_from_file_location("art_report", path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module.report_for_workbook
    finally:
        sys.path.remove(str(path.parent))


def run_scenario(copy: Path, scenario: Scenario, port: int, *, serve=MockJira, session=com.excel_session, export=None, status=status_of,
                 report=None) -> int:
    """The scenario alone, unattended, in one process: the fake Jira of the scenario (this process, nothing else listens), the copy
    opened in a hidden Excel with its macros, the extraction (modActions.ExtractAndCompute, synchronous), the charts exported to
    out/manual_test/charts/<scenario> (SVG, PNG, index.html), the copy saved WITH its results. Excel and the fake Jira are
    always closed. The hooks (serve, session, export, status) let the tests run it without Excel."""
    if export is None:
        from jira2dashboard.flow import sheets               # the COM side: only the unattended run needs it
        export = sheets.export_charts_to_look_at
    jira = serve(scenario.dataset, port=port, credentials=(MOCK_EMAIL, MOCK_TOKEN)).start()
    try:
        with session() as app:
            app.AutomationSecurity = com.MSO_AUTOMATION_SECURITY_LOW
            wb = app.Workbooks.Open(str(copy.resolve()))
            try:
                app.Run(f"'{wb.Name}'!modActions.ExtractAndCompute")
                print(f"mock-workbook: extraction done: {status(wb)}")
                text = ""
                if report is not None and scenario.name == "art-metrics":      # the only scenario whose numbers are worked out by hand
                    text = report(wb)
                export(wb, scenario.name.replace("-", "_"), text)
                wb.Save()
            finally:
                wb.Close(False)
    finally:
        jira.stop()
    return 0


def port_of(args, scenario: Scenario | None) -> int:
    """The port asked for, else the one of the scenario (its own fake Jira instance), else 8099."""
    return args.port if args.port is not None else (8099 if scenario is None else scenario.port)


def fresh_folder(folder: Path, keep: Path) -> None:
    """The folder of a scenario before a run: the results of the previous run (xlsx of the Python path, old CSV) go, the logs stay."""
    for old in folder.iterdir() if folder.is_dir() else ():
        if old.is_file() and old != keep:
            try:
                old.unlink()
            except OSError:
                pass                                              # a file Excel still holds: it is rewritten or ignored


def main(argv: list[str] | None = None) -> int:
    args = parse(argv)
    source = Path(args.workbook)
    if not source.is_file():
        print(f"mock-workbook: no workbook {source}: run scripts\\run.ps1 build first")
        return 1
    if args.run and not args.scenario:
        print("mock-workbook: --run needs --scenario (the unattended run is the one of a scenario)")
        return 2
    scenario = scenarios.get(args.scenario) if args.scenario else None
    copy = copy_path(scenario)
    copy.parent.mkdir(parents=True, exist_ok=True)
    if scenario is not None:
        fresh_folder(copy.parent, copy)
    try:
        shutil.copyfile(source, copy)
    except PermissionError:
        print(f"mock-workbook: close {copy.name} first (it is still open in Excel)")
        return 1
    port = port_of(args, scenario)
    url = base_url(port)
    fill_copy(copy, url) if scenario is None else fill_copy(copy, url, scenario)
    time.sleep(SETTLE_S)
    if not holds_the_fill(copy, url):
        print(f"mock-workbook: {com.lost_message(copy, 'the address of the fake Jira in the sheet Config')}")
        return 1
    print(f"mock-workbook: {copy}")
    csv_text = rovo_csv(port) if scenario is None else rovo_csv(port, scenario)
    csv_file = csv_path(scenario)
    csv_file.write_bytes(csv_text.encode("utf-8"))
    print(f"mock-workbook: {csv_file} (to try Load settings CSV; type the connection by hand)")
    if args.run:
        run_scenario(copy, scenario, port, report=report_builder())
        print(f"mock-workbook: the copy holds its results: {copy}")
        print(f"mock-workbook: the charts: {CHARTS_DIR / scenario.name.replace('-', '_') / 'index.html'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
