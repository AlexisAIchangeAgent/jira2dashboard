"""The copy of a workbook for a manual test against the fake Jira: what goes into Config, Periods and PIs."""
import datetime as dt

from jira2dashboard import clean
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
    periods = M.iterations()
    assert periods[0].name == "I1" and periods[0].start == SINCE
    assert all(p.end - p.start == dt.timedelta(days=14) for p in periods)
    assert all(a.end == b.start for a, b in zip(periods, periods[1:]))
    assert periods[-1].contains(AS_OF) and not periods[-2].contains(AS_OF)
    assert [p.name for p in periods] == [f"I{n}" for n in range(1, len(periods) + 1)]


def test_the_pis_sheet_gets_the_pi_of_the_known_anomalies():
    assert [(p.name, p.start, p.end) for p in M.PIS] == [(PI, *PI_WINDOW)]


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
