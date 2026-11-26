"""The command `run.ps1 extract`: its log, its exit codes, the secrets it never writes (fake Jira in process)."""
import base64
import re

import pytest

from jira2dashboard import mock_workbook as M
from jira2dashboard.mockjira.generator import generate
from jira2dashboard.mockjira.server import MOCK_EMAIL, MOCK_TOKEN, MockJira
from jira2dashboard.params.model import project_keys
from jira2dashboard.standalone import __main__ as X
from jira2dashboard.standalone import xlsx_write
from jira2dashboard.standalone.proxy import Detected


def direct(url):
    return Detected("", "", "direct (test)")


@pytest.fixture(scope="module")
def dataset():
    return generate(seed=7, issues=30, projects=project_keys())


@pytest.fixture
def jira(dataset):
    with MockJira(dataset, credentials=(MOCK_EMAIL, MOCK_TOKEN)) as server:
        yield server


def run(tmp_path, workbook, log_name="python.log", **kwargs) -> tuple[int, str]:
    log = tmp_path / log_name
    code = X.main(["--workbook", str(workbook), "--out", str(tmp_path / "out.xlsx"), "--log", str(log)]
                  + kwargs.pop("extra", []), sleep=lambda s: None, detect=direct, **kwargs)
    return code, log.read_text(encoding="utf-8") if log.exists() else ""


def params(tmp_path, url):
    return M.write_params_book(tmp_path / "params.xlsx", url, {"ProxyServer": "DIRECT", "MaxRetries": "1"})


def test_a_secret_that_jira_quotes_back_is_masked_in_the_log_and_on_the_console(tmp_path, jira, capsys):
    jira.fail_next("GET", "/rest/api/3/search/jql", 400,
                   body={"errorMessages": [f"token {MOCK_TOKEN} refused for Basic {base64.b64encode(b'x').decode()}"]})
    code, log = run(tmp_path, params(tmp_path, jira.base_url))
    console = capsys.readouterr().out
    assert code == 1 and f"token {X.MASK} refused" in log
    assert MOCK_TOKEN not in log and MOCK_TOKEN not in console


def test_verbose_logs_each_call_that_worked(tmp_path, jira):
    code, log = run(tmp_path, params(tmp_path, jira.base_url), extra=["--verbose"])
    assert code == 0 and "search: GET /rest/api/3/search/jql -> HTTP 200" in log
    _code, quiet = run(tmp_path, params(tmp_path, jira.base_url), log_name="quiet.log")
    assert "-> HTTP 200" not in quiet


def test_the_log_gives_the_inputs_and_the_time_of_each_step_and_of_the_run(tmp_path, jira):
    """To debug from the log alone: a PARAM line per value of Config (never the token), the size of the other parameter
    sheets, the time at the end of each END line and of the line RUN end."""
    code, log = run(tmp_path, params(tmp_path, jira.base_url))
    assert code == 0
    assert log.splitlines()[0].endswith(" | jira2dashboard unknown")              # a workbook without J2D_Version
    assert "PARAM ProxyServer = DIRECT\n" in log and "PARAM MaxRetries = 1\n" in log and "PARAM Filter = (empty)\n" in log
    assert "PARAM ApiToken = (set)\n" in log and MOCK_TOKEN not in log
    assert "PARAM sheet Projects = " in log and "PARAM sheet Iterations = " in log
    timed = re.compile(r" \| \d+\.\d{3} s$")
    ends = [line for line in log.splitlines() if " END " in line or " RUN end " in line]
    assert len(ends) >= 6 and all(timed.search(line) for line in ends), ends


def test_a_stopped_run_gives_its_code_and_its_time_on_its_last_line(tmp_path):
    code, log = run(tmp_path, tmp_path / "missing.xlsm")
    assert code == 1
    assert re.search(r"RUN end stopped J2D-SYS-011 \(log: .*\) \| \d+\.\d{3} s$", log.splitlines()[-1])


def test_a_wrong_account_stops_before_the_search_although_jira_cloud_would_answer_empty_lists(tmp_path, dataset):
    """AL, 2026-11-26: Jira Cloud serves a wrong ApiToken as an unknown visitor (200, no ticket); only myself says 401. Without the check the
    run would write an empty workbook with exit code 0."""
    with MockJira(dataset, credentials=(MOCK_EMAIL, MOCK_TOKEN), anonymous=True) as server:
        book = M.write_params_book(tmp_path / "params.xlsx", server.base_url, {"ProxyServer": "DIRECT", "MaxRetries": "1",
                                                                              "ApiToken": "a-wrong-token"})
        code, log = run(tmp_path, book)
        asked = [r.path for r in server.requests]
    assert code == 1 and asked == ["/rest/api/3/myself"], asked
    assert "J2D-JIRA-401" in log and "In the sheet Config_from_user check Email and ApiToken" in log
    assert "BEGIN login" in log and "END login" not in log                     # the step that stopped is named, as in the VBA log
    assert not (tmp_path / "out.xlsx").exists()


def test_parameter_lines_never_show_the_token():
    config = xlsx_write.Sheet("Config", ("Parameter", "Value"), [("ApiToken", "secret"), ("Email", "a@b.c"), ("Filter", None)])
    lines = X.parameter_lines([config, xlsx_write.Sheet("Projects", ("Key",), [("A",), ("B",)]),
                               xlsx_write.Sheet("PIs", ("PI",), [("PI#01-2026",)]), xlsx_write.Sheet("Iterations", ("Name",), [])])
    assert lines == ["ApiToken = (set)", "Email = a@b.c", "Filter = (empty)", "sheet Projects = 2 rows", "sheet PIs = 1 row",
                     "sheet Iterations = 0 rows"]


@pytest.mark.parametrize("content, message", [(None, "No such file"), (b"not a zip", "File is not a zip file")])
def test_a_workbook_that_is_missing_or_not_a_workbook_is_named(tmp_path, content, message):
    workbook = tmp_path / "jira2dashboard.xlsm"
    if content is not None:
        workbook.write_bytes(content)
    code, log = run(tmp_path, workbook)
    assert code == 1 and message in log and "RUN end stopped" in log


def test_results_open_in_excel_are_left_alone_with_a_clear_message(tmp_path, jira, monkeypatch):
    workbook = params(tmp_path, jira.base_url)

    def locked(*_args):
        raise PermissionError(13, "in use")
    monkeypatch.setattr(xlsx_write.os, "replace", locked)
    code, log = run(tmp_path, workbook)
    assert code == 1 and "close out.xlsx first" in log


def test_an_unexpected_error_leaves_its_traceback_in_the_log(tmp_path, jira):
    class Broken:
        def __init__(self, *args):
            pass

        def send(self, *args):
            raise RuntimeError("the transport broke")
    code, log = run(tmp_path, params(tmp_path, jira.base_url), transport=Broken)
    assert code == 1 and "unexpected error" in log and "Traceback" in log and "RuntimeError: the transport broke" in log


def test_the_default_output_is_next_to_the_workbook(tmp_path, jira):
    workbook = params(tmp_path, jira.base_url)
    log = tmp_path / "python.log"
    assert X.main(["--workbook", str(workbook), "--log", str(log)], sleep=lambda s: None, detect=direct) == 0
    assert (tmp_path / X.OUT_NAME).exists()


def test_the_default_log_is_the_folder_logs_next_to_the_workbook(tmp_path, jira):
    workbook = params(tmp_path, jira.base_url)
    assert X.main(["--workbook", str(workbook)], sleep=lambda s: None, detect=direct) == 0
    assert "RUN" in (tmp_path / "logs" / "python.log").read_text(encoding="utf-8")


@pytest.mark.parametrize("date, line", [
    (("created", "ms"), "changelog date key: created (epoch ms)"),                      # what the real Jira sends
    (("created", "iso"), "changelog date key: created (ISO)"),                          # what the OpenAPI says
    (("createdAt", "ms"), "changelog date key: createdAt (epoch ms)"),                  # a label Atlassian could change
    (("timestamp", "iso"), "changelog date key: timestamp (ISO)"),
])
def test_the_date_of_a_changelog_is_found_by_its_value_and_named_in_the_log(tmp_path, dataset, date, line):
    with MockJira(dataset, credentials=(MOCK_EMAIL, MOCK_TOKEN), changelog_date=date) as server:
        code, log = run(tmp_path, params(tmp_path, server.base_url))
    assert code == 0 and line in log and "J2D-DATA-003" not in log


# --- the type of every column written from Jira, and what Jira really sent (the log lines that show a v0.7.0 failure at once) ----------------

def test_the_log_says_what_jira_sent_and_that_every_column_has_its_type(tmp_path, jira):
    code, log = run(tmp_path, params(tmp_path, jira.base_url))
    assert code == 0
    assert "wire types issue | id string, key string, fields.created string, fields.resolutiondate string, fields.status.name string" in log
    assert "wire types configured fields | WorkTypeField string, SprintField array, StoryPointsField number" in log
    assert "wire types changelog | id string, created number, fieldId string, from string, to string" in log     # Jira sends epoch ms
    assert re.search(r"column types \| tables \d \| columns \d+ \| checked \d+ \| problems 0\b", log)
    assert "types Issues | rows 30 | filled: Key 30, Id 30" in log and "types Changelog | rows " in log
    assert not re.search(r"J2D-DATA-0(0[4-9]|1[0-3])", log)


def test_the_wire_types_follow_a_renamed_date_key(tmp_path, dataset):
    with MockJira(dataset, credentials=(MOCK_EMAIL, MOCK_TOKEN), changelog_date=("createdAt", "iso")) as server:
        code, log = run(tmp_path, params(tmp_path, server.base_url))
    assert code == 0 and "wire types changelog | id string, createdAt string, fieldId string" in log
    assert "problems 0" in log


def test_a_wrong_type_is_a_warning_with_its_code_and_two_info_lines_and_stops_nothing(caplog):
    import logging
    import types
    row = ("ABC-1", "10001", "5", 1761134538011, "1761134538011", "status", "Status", "jira", "1", "To Do", "2", "Done")
    caplog.set_level(logging.INFO, logger="jira2dashboard.extract")
    report = X.log_column_types(types.SimpleNamespace(rows=[], links=[], changelog=[row], flow=[]), [])
    assert [p.name for p in report.problems] == ["Changelog.CreatedUtc"]
    records = [(r.levelname, r.getMessage()) for r in caplog.records]
    assert ("WARNING", "J2D-DATA-004 | Changelog.CreatedUtc | expected datetime | found number | source histories[].created | "
                       "first filled D2") in records
    assert ("INFO", "J2D-DATA-004 | Changelog.CreatedUtc | rows 1, filled 1, empty 0 | kinds in first 20: number x 1 | "
                    "value shape Double:9999999999999") in records
    assert any(level == "INFO" and text.startswith("J2D-DATA-004 | Changelog.CreatedUtc | hint: 13 digits: epoch milliseconds")
               for level, text in records)
    assert ("INFO", "column types | tables 1 | columns 12 | checked 12 | problems 1") in records
    order = [text for _level, text in records]
    assert order.index(next(t for t in order if t.startswith("J2D-DATA-004"))) < order.index(next(t for t in order if t.startswith("types Changelog")))


# --- what the status rules hold, and why a metric cannot be computed (AL, 2026-11-22) -------------------------------------------------------

def test_the_log_counts_the_status_rules_of_each_workflow_kind(tmp_path, jira):
    code, log = run(tmp_path, params(tmp_path, jira.base_url))
    assert code == 0
    assert re.search(r"INFO .*status rules \| Feature: start [1-9]\d*, active \d+, wait \d+, end \d+ \| Team: start [1-9]", log), log
    assert "J2D-FLOW-010" not in log and "J2D-FLOW-011" not in log


def test_without_the_project_administrator_right_the_workflows_are_not_asked(tmp_path, jira):
    book = M.write_params_book(tmp_path / "params.xlsx", jira.base_url, {"ProxyServer": "DIRECT", "MaxRetries": "1", "IsUserProjectAdmin": "False"})
    code, log = run(tmp_path, book)
    assert code == 0
    assert "workflows of Jira not read: IsUserProjectAdmin is False" in log
    assert not [r for r in jira.requests if r.path.endswith("/workflows/preview")]


def test_the_log_and_the_sheet_run_say_what_to_look_at(tmp_path, jira):
    """The fields of modHealth (flow/health.py) in the log and in the sheet Run: none for the clean fake data, AsOf before the data when AsOf is
    set before its last status changes (2026-11-01)."""
    from jira2dashboard.standalone.xlsx_read import read_sheets
    code, log = run(tmp_path, params(tmp_path, jira.base_url))
    assert code == 0 and " health | " not in log
    assert ["Health", "nothing to look at"] in read_sheets(tmp_path / "out.xlsx", {X.RUN_SHEET})[X.RUN_SHEET]
    book = M.write_params_book(tmp_path / "early.xlsx", jira.base_url, {"ProxyServer": "DIRECT", "MaxRetries": "1", "AsOf": "2026-10-15 00:00"})
    code, log = run(tmp_path, book, log_name="early.log")
    assert code == 0
    assert re.search(r"INFO health \| AsOf before the data: \d+ status changes after AsOf \(Flow Load stops at AsOf\)", log), log
    health = next(row[1] for row in read_sheets(tmp_path / "out.xlsx", {X.RUN_SHEET})[X.RUN_SHEET] if row and row[0] == "Health")
    assert "AsOf before the data: " in health


def test_a_refused_workflow_reading_is_a_warning_with_its_code(tmp_path, jira):
    jira.fail_next("POST", "/rest/api/3/workflows/preview", 403)
    code, log = run(tmp_path, params(tmp_path, jira.base_url))
    assert code == 0 and "J2D-JIRA-030 | the workflows of Jira could not be read" in log


def test_the_xlsx_holds_the_jira_sheets_in_the_layout_of_the_vba_and_the_rules_made_of_them(tmp_path, jira, dataset):
    """Step F: what the run read of Jira's configuration is kept (a later compute reads it again): the four sheets read back give the workflows of the
    fake Jira, the catalog its statuses; StatusRules holds the rules with a start for each workflow kind."""
    from jira2dashboard import sheet_names as S
    from jira2dashboard.flow import jira_workflows as JW
    from jira2dashboard.mockjira.flow_adapter import jira_workflows
    from jira2dashboard.standalone.xlsx_read import read_sheets
    code, _log = run(tmp_path, params(tmp_path, jira.base_url))
    assert code == 0
    names = {S.of(n): n for n in ("JiraStatusCatalog", "JiraIssueTypes", "JiraWorkflows", "JiraTransitions", "StatusRules")}
    sheets = {names[k]: rows for k, rows in read_sheets(tmp_path / "out.xlsx", set(names)).items()}
    assert set(sheets) == set(names.values()) and tuple(sheets["JiraWorkflows"][0]) == JW.WORKFLOWS_HEADER
    info = JW.from_sheets(*(sheets[n][1:] for n in ("JiraIssueTypes", "JiraWorkflows", "JiraTransitions")))
    expected = jira_workflows(dataset)
    assert (info.moves, info.starts, info.initial) == (expected.moves, expected.starts, expected.initial) and info.workflow_of
    assert len(JW.catalog_of(sheets["JiraStatusCatalog"][1:])) > 5
    rules = sheets["StatusRules"][1:]
    assert {r[0] for r in rules if r[3] == 1} == {"Feature", "Team"} and {r[6] for r in rules} <= {"Jira", "user"}
