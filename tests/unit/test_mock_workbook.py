"""The copy of a workbook for a manual test against the fake Jira: what goes into Config, Iterations and PIs."""
import csv
import datetime as dt
import io

from jira2dashboard import clean
from jira2dashboard import rovo_prompt as RP
from jira2dashboard import mock_workbook as M
from jira2dashboard import sheet_names as S
from jira2dashboard.mockjira.anomalies import PI, PI_WINDOW
from jira2dashboard.mockjira.dataset import FieldIds
from jira2dashboard.mockjira.generator import AS_OF, SINCE, generate
from jira2dashboard.mockjira.server import MOCK_EMAIL, MOCK_TOKEN
from jira2dashboard.params.model import load_all
from jira2dashboard.paths import OUT_DIR


def test_the_config_points_to_the_fake_jira_on_the_port_with_its_account_and_reference_date():
    values = M.config_values(M.base_url(8099))
    assert {k: values[k] for k in ("BaseUrl", "Email", "ApiToken", "AsOf")} == {
        "BaseUrl": "http://127.0.0.1:8099", "Email": MOCK_EMAIL, "ApiToken": MOCK_TOKEN, "AsOf": "2026-11-03 18:00"}


def test_the_rovo_csv_gives_the_configuration_of_the_fake_jira_but_never_the_connection():
    text = M.rovo_csv()
    assert text.startswith("\ufeffSheet,Row,Column,Value,Note\r\n")
    rows = list(csv.reader(io.StringIO(text.removeprefix("\ufeff"))))[1:]
    config = {column: value for sheet, _row, column, value, _note in rows if sheet == S.of("Config")}
    assert config["AsOf"] == f"{AS_OF:%Y-%m-%d %H:%M}" and config["StoryPointsField"] == FieldIds().story_points
    assert not set(config) & set(RP.PROTECTED)                                  # the button never loads them anyway
    cells = {(sheet, int(row), column): value for sheet, row, column, value, _note in rows if sheet != S.of("Config")}
    first = M.iterations()[0]
    assert cells[(S.of("Iterations"), 1, "Iteration-label")] == first.name and cells[(S.of("Iterations"), 1, "Start")] == f"{first.start:%Y-%m-%d}"
    assert len([1 for sheet, row, column in cells if sheet == S.of("Iterations") and column == "End"]) == len(M.iterations())
    assert cells[(S.of("PIs"), 1, "PI")] == PI
    assert {S.logical(sheet) for sheet, _row, _column in cells} == set(RP.SHEET_RULES)
    projects = next(sheet for sheet in load_all() if sheet.name == "Projects")
    assert len([1 for sheet, _row, column in cells if sheet == S.of("Projects") and column == "Project"]) == len(projects.rows)


def test_a_date_reads_as_rovo_writes_it():
    assert M.csv_value(dt.datetime(2026, 10, 5)) == "2026-10-05"
    assert M.csv_value(dt.datetime(2026, 11, 3, 18, 0)) == "2026-11-03 18:00"
    assert M.csv_value(None) == "" and M.csv_value(3) == "3"


def test_the_reference_date_is_the_one_of_the_generated_data():
    assert generate(issues=5).as_of == AS_OF
    assert M.config_values("u")["AsOf"] == f"{AS_OF:%Y-%m-%d %H:%M}"


def test_the_config_names_every_custom_field_of_the_fake_jira():
    values, fields = M.config_values("u"), FieldIds()
    assert values["WorkTypeField"] == fields.work_type
    assert {name: values[name] for name in M.FIELD_IDS} == {
        "StoryPointsField": fields.story_points, "BusinessValueField": fields.business_value, "PIField": fields.pi,
        "TeamField": fields.team, "SprintField": fields.sprint, "TargetStartField": fields.target_start,
        "TargetEndField": fields.target_end}


def test_every_name_the_copy_writes_is_a_parameter_of_the_config_sheet():
    config = next(sheet for sheet in load_all() if sheet.name == "Config")
    assert set(M.config_values("u")) <= {row[0] for row in config.rows}


def test_the_iterations_last_14_days_follow_each_other_from_the_first_issue_and_the_last_holds_the_reference_date():
    iterations = M.iterations()
    assert iterations[0].name == "I1" and iterations[0].start == SINCE
    assert all(p.end - p.start == dt.timedelta(days=14) for p in iterations)
    assert all(a.end == b.start for a, b in zip(iterations, iterations[1:]))
    assert iterations[-1].contains(AS_OF) and not iterations[-2].contains(AS_OF)
    assert [p.name for p in iterations] == [f"I{n}" for n in range(1, len(iterations) + 1)]


def test_the_pis_sheet_gets_the_pi_of_the_known_anomalies_then_the_other_pi_of_the_generator():
    assert [(p.name, p.start, p.end) for p in M.PIS] == [(PI, *PI_WINDOW), ("PI#02-2026", *PI_WINDOW)]


def test_the_copy_is_jira2dashboard4manualtest_in_out_manual_test_which_clean_never_empties(tmp_path):
    assert (M.COPY.parent.name, M.COPY.name) == ("manual_test", "jira2dashboard4manualtest.xlsm")
    copy = tmp_path / M.COPY.relative_to(OUT_DIR)
    copy.parent.mkdir(parents=True)
    copy.write_bytes(b"x")
    clean.clean(tmp_path)
    assert copy.exists()


def test_a_missing_workbook_is_refused(tmp_path, capsys):
    assert M.main(["--workbook", str(tmp_path / "none.xlsm")]) == 1
    assert "no workbook" in capsys.readouterr().out


def test_a_copy_still_open_in_excel_is_named_not_a_raw_error(tmp_path, capsys, monkeypatch):
    source = tmp_path / "jira2dashboard.xlsm"
    source.write_bytes(b"x")
    monkeypatch.setattr(M, "COPY", tmp_path / "inter" / M.COPY_NAME)

    def locked(*_args):
        raise PermissionError(13, "The process cannot access the file")
    monkeypatch.setattr(M.shutil, "copyfile", locked)
    assert M.main(["--workbook", str(source)]) == 1
    assert f"close {M.COPY_NAME} first" in capsys.readouterr().out


# --- Excel's writes undone (an endpoint protection), seen on another machine: the copy stayed empty ------------------------

def prepare_main(tmp_path, monkeypatch, written):
    """`main` with a source, a copy in tmp_path and a fill_copy that leaves `written` (the Config values it saves)."""
    from jira2dashboard.standalone.xlsx_write import write_xlsx
    source = tmp_path / "jira2dashboard.xlsm"
    source.write_bytes(b"x")
    monkeypatch.setattr(M, "COPY", tmp_path / "out" / M.COPY_NAME)
    monkeypatch.setattr(M, "SETTLE_S", 0)
    monkeypatch.setattr(M, "fill_copy", lambda path, url: write_xlsx(path, M.params_book(url, written)))
    return source


def test_a_copy_that_holds_the_fill_is_accepted(tmp_path, monkeypatch, capsys):
    source = prepare_main(tmp_path, monkeypatch, None)
    assert M.main(["--workbook", str(source)]) == 0
    assert f"mock-workbook: {M.COPY}" in capsys.readouterr().out


def test_a_copy_that_does_not_hold_the_fill_is_named_with_its_code_not_left_for_the_extraction(tmp_path, monkeypatch, capsys):
    source = prepare_main(tmp_path, monkeypatch, {"BaseUrl": ""})
    assert M.main(["--workbook", str(source)]) == 1
    out = capsys.readouterr().out
    assert "J2D-SYS-013" in out and "endpoint protection" in out and "check-machine" in out
    assert not (M.COPY.parent / M.CSV_NAME).exists()


def test_the_messages_of_the_com_layer_carry_codes_of_the_table():
    from jira2dashboard import error_codes
    from jira2dashboard.excel import com
    from pathlib import Path
    codes = {row.code for row in error_codes.TABLE}
    assert {com.MISSING_CODE, com.LOST_CODE} <= codes
    target = Path("C:/x/out/release/jira2dashboard.xlsm")
    assert com.MISSING_CODE in com.missing_message(target, "after the save") and "jira2dashboard.xlsm" in com.missing_message(target, "x")
    assert com.LOST_CODE in com.lost_message(target, "the change")


# --- a scenario is a separate instance: its own port, its own folder, its own copy (AL, 2026-11-17) ------------------------

def test_a_scenario_has_its_own_port_folder_and_copy_so_the_generic_manual_test_is_left_alone():
    from jira2dashboard.mockjira.scenarios import get
    scenario = get("art-metrics")
    assert scenario.port == 8101 != 8099
    assert M.copy_path(None) == M.COPY and M.csv_path(None) == M.COPY.parent / M.CSV_NAME
    copy = M.copy_path(scenario)
    assert (copy.parent.parent, copy.parent.name, copy.name) == (M.COPY.parent, "art-metrics", "jira2dashboard4art-metrics.xlsm")
    assert M.csv_path(scenario) == copy.parent / "rovo_art-metrics.csv"


def test_the_scripts_use_the_port_of_the_scenario():
    from pathlib import Path
    from jira2dashboard.mockjira.scenarios import get
    scripts = Path(__file__).resolve().parents[2] / "scripts"
    for name in ("manual-test.ps1", "start-mock.ps1"):
        assert f"$ScenarioPort = {get('art-metrics').port}" in (scripts / name).read_text(encoding="utf-8"), name


def test_fill_workbook_writes_the_parameter_sheets_of_a_scenario_Projects_included(monkeypatch):
    from jira2dashboard.mockjira.scenarios import get
    scenario, calls = get("art-metrics"), []

    class Book:
        def Worksheets(self, name):
            return name
    for fn in ("write_config", "write_iterations", "write_rows"):
        monkeypatch.setattr(M, fn, lambda ws, data, fn=fn: calls.append((fn, ws, data)))
    M.fill_workbook(Book(), "http://x", scenario)
    written = {ws: data for fn, ws, data in calls}
    assert [fn for fn, _ws, _d in calls].count("write_rows") == 3
    assert written[S.of("Projects")] == (("TTT", "Program", None),)
    assert {r[0] for r in written[S.of("Workflows")]} == {"Feature"} and not written[S.of("Transitions")]          # exceptions only (the Wait statuses)
    assert len(written[S.of("Iterations")]) == 24 and [p.name for p in written[S.of("PIs")]] == ["PI-1", "PI-2", "PI-3", "PI-4"]


def test_fill_workbook_without_a_scenario_writes_config_iterations_and_pis_only(monkeypatch):
    calls = []

    class Book:
        def Worksheets(self, name):
            return name
    for fn in ("write_config", "write_iterations", "write_rows"):
        monkeypatch.setattr(M, fn, lambda ws, data, fn=fn: calls.append(fn))
    M.fill_workbook(Book(), "http://x")
    assert calls == ["write_config", "write_iterations", "write_iterations"]


def test_main_writes_the_copy_of_a_scenario_in_its_own_folder(tmp_path, monkeypatch, capsys):
    from jira2dashboard.standalone.xlsx_write import write_xlsx
    source = tmp_path / "jira2dashboard.xlsm"
    source.write_bytes(b"x")
    monkeypatch.setattr(M, "COPY", tmp_path / "out" / M.COPY_NAME)
    monkeypatch.setattr(M, "SETTLE_S", 0)
    seen = {}

    def fill(path, url, scenario=None):
        seen["url"] = url
        write_xlsx(path, M.params_book(url, scenario=scenario))
    monkeypatch.setattr(M, "fill_copy", fill)
    assert M.main(["--workbook", str(source), "--scenario", "art-metrics"]) == 0
    folder = tmp_path / "out" / "art-metrics"
    assert (folder / "jira2dashboard4art-metrics.xlsm").is_file() and (folder / "rovo_art-metrics.csv").is_file()
    assert seen["url"] == "http://127.0.0.1:8101"                                 # the port of the scenario, not 8099
    assert not (tmp_path / "out" / M.COPY_NAME).exists()                          # the generic copy is not touched
    assert "rovo_art-metrics.csv" in capsys.readouterr().out


# --- --run: the scenario alone, unattended: serve it, extract with the macros of the copy, export the charts, save, close ------

class FakeJira:
    def __init__(self, calls):
        self.calls = calls

    def start(self):
        self.calls.append("jira start")
        return self

    def stop(self):
        self.calls.append("jira stop")


class FakeBook:
    Name = "copy.xlsm"

    def __init__(self, calls, fail=False):
        self.calls, self.fail = calls, fail

    def Save(self):
        self.calls.append("save")

    def Close(self, save):
        self.calls.append(f"close save={save}")

    class Names:
        pass


def fake_session(calls, fail=False):
    import contextlib

    class App:
        AutomationSecurity = 0

        class Workbooks:
            @staticmethod
            def Open(path):
                calls.append(f"open {path.rsplit(chr(92), 1)[-1]}")
                return FakeBook(calls)

        @staticmethod
        def Run(macro):
            calls.append(f"run {macro}")
            if fail:
                raise RuntimeError("macro failed")

    @contextlib.contextmanager
    def session():
        calls.append("excel start")
        try:
            yield App()
        finally:
            calls.append("excel quit")
    return session


def run_with_fakes(tmp_path, fail=False):
    from jira2dashboard.mockjira.scenarios import get
    calls, exported = [], []
    copy = tmp_path / "copy.xlsm"
    copy.write_bytes(b"x")
    result = None
    try:
        result = M.run_scenario(copy, get("art-metrics"), 8101, serve=lambda ds, **o: FakeJira(calls),
                                session=fake_session(calls, fail), export=lambda wb, name, report="": exported.append((name, report)) or calls.append("export"),
                                status=lambda wb: "Done | Issues: 25", report=lambda wb: "REPORT")
    except RuntimeError:
        pass
    return result, calls, exported


def test_run_serves_the_scenario_extracts_exports_saves_and_closes_in_that_order(tmp_path):
    result, calls, exported = run_with_fakes(tmp_path)
    assert result == 0 and exported == [("art_metrics", "REPORT")]                                   # the charts folder: charts\art_metrics
    assert calls == ["jira start", "excel start", "open copy.xlsm", "run 'copy.xlsm'!modActions.ExtractAndCompute", "export", "save",
                     "close save=False", "excel quit", "jira stop"]


def test_run_stops_the_fake_jira_and_excel_even_when_the_macro_fails(tmp_path):
    result, calls, exported = run_with_fakes(tmp_path, fail=True)
    assert result is None and exported == []                                             # the error reached the caller
    assert calls[-3:] == ["close save=False", "excel quit", "jira stop"] and "save" not in calls


def test_run_needs_a_scenario(tmp_path, capsys):
    source = tmp_path / "jira2dashboard.xlsm"
    source.write_bytes(b"x")
    assert M.main(["--workbook", str(source), "--run"]) == 2
    assert "--run needs --scenario" in capsys.readouterr().out


def test_the_status_of_a_run_counts_the_issues_and_gives_the_headline_and_never_raises():
    class Rows:
        Count = 25

    class Table:
        ListRows = Rows

    class Sheet:
        @staticmethod
        def ListObjects(name):
            return Table

    class Cell:
        RefersToRange = type("R", (), {"Value": "Ready"})

    class Book:
        @staticmethod
        def Worksheets(name):
            return Sheet

        @staticmethod
        def Names(name):
            return Cell
    assert M.status_of(Book()) == f"25 issue(s) in the sheet {S.of('Issues')}, status Ready"
    assert M.status_of(object()) == "?"


def test_a_scenario_folder_loses_the_results_of_the_last_run_but_keeps_its_logs(tmp_path):
    (tmp_path / "logs").mkdir()
    (tmp_path / "logs" / "jira2dashboard.log").write_text("RUN")
    (tmp_path / "jira2dashboard.xlsx").write_text("old results")
    (tmp_path / "rovo_art-metrics.csv").write_text("old")
    keep = tmp_path / "jira2dashboard4art-metrics.xlsm"
    keep.write_text("copy")
    M.fresh_folder(tmp_path, keep)
    assert sorted(p.name for p in tmp_path.iterdir()) == ["jira2dashboard4art-metrics.xlsm", "logs"]
    assert (tmp_path / "logs" / "jira2dashboard.log").exists()
    M.fresh_folder(tmp_path / "missing", keep)                                    # no folder yet: nothing to do


def test_the_report_builder_is_the_one_of_tests_art_report_in_a_checkout_and_none_without_tests(tmp_path, monkeypatch):
    builder = M.report_builder()
    assert callable(builder) and builder.__name__ == "report_for_workbook"
    monkeypatch.setattr(M, "REPO_ROOT", tmp_path)                                 # no tests folder: the report is left out
    assert M.report_builder() is None
