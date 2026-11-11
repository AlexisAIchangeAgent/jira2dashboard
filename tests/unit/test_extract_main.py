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


def test_parameter_lines_never_show_the_token():
    config = xlsx_write.Sheet("Config", ("Parameter", "Value"), [("ApiToken", "secret"), ("Email", "a@b.c"), ("Filter", None)])
    lines = X.parameter_lines([config, xlsx_write.Sheet("Projects", ("Key",), [("A",), ("B",)])])
    assert lines == ["ApiToken = (set)", "Email = a@b.c", "Filter = (empty)", "sheet Projects = 2 rows"]


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
