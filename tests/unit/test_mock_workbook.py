"""The copy of a workbook for a manual test against the fake Jira: what goes into Config, Iterations and PIs."""
import csv
import datetime as dt
import io

from jira2dashboard import clean
from jira2dashboard import rovo_prompt as RP
from jira2dashboard import mock_workbook as M
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
    config = {column: value for sheet, _row, column, value, _note in rows if sheet == "Config"}
    assert config["AsOf"] == f"{AS_OF:%Y-%m-%d %H:%M}" and config["StoryPointsField"] == FieldIds().story_points
    assert not set(config) & set(RP.PROTECTED)                                  # the button never loads them anyway
    cells = {(sheet, int(row), column): value for sheet, row, column, value, _note in rows if sheet != "Config"}
    first = M.iterations()[0]
    assert cells[("Iterations", 1, "Iteration-label")] == first.name and cells[("Iterations", 1, "Start")] == f"{first.start:%Y-%m-%d}"
    assert len([1 for sheet, row, column in cells if sheet == "Iterations" and column == "End"]) == len(M.iterations())
    assert cells[("PIs", 1, "PI")] == PI
    assert {sheet for sheet, _row, _column in cells} == set(RP.SHEET_RULES)
    projects = next(sheet for sheet in load_all() if sheet.name == "Projects")
    assert len([1 for sheet, _row, column in cells if sheet == "Projects" and column == "Project"]) == len(projects.rows)


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
