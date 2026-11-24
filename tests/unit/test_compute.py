"""`run.ps1 compute` (step F): the results made again from the sheets of a filled workbook, without Jira. The oracle needs no Excel: extract with
Python from the fake Jira, compute from that xlsx, and every sheet but Run and Errors must be the same; a corrected sheet Workflows must give what
a fresh extraction with the same exception gives."""
import pytest

from jira2dashboard import mock_workbook as M
from jira2dashboard import sheet_names as S
from jira2dashboard.mockjira.generator import generate
from jira2dashboard.mockjira.scenarios import get
from jira2dashboard.mockjira.server import MOCK_EMAIL, MOCK_TOKEN, MockJira
from jira2dashboard.params.model import ParamsError, project_keys
from jira2dashboard.standalone import __main__ as X
from jira2dashboard.standalone import recompute
from jira2dashboard.standalone.proxy import Detected
from jira2dashboard.standalone.xlsx_read import read_sheets
from jira2dashboard.standalone.xlsx_write import Sheet, write_xlsx

LEFT_OUT = {X.RUN_SHEET, X.ERRORS_SHEET}               # the time, the paths and the log of the run


def direct(url):
    return Detected("", "", "direct (test)")


class NoJira:
    """A transport that fails: compute must never call Jira."""

    def __init__(self, *args):
        raise AssertionError("compute called Jira")


def extract(tmp_path, book, out="extracted.xlsx"):
    code = X.main(["--workbook", str(book), "--out", str(tmp_path / out), "--log", str(tmp_path / "extract.log")], sleep=lambda s: None, detect=direct)
    assert code == 0, (tmp_path / "extract.log").read_text(encoding="utf-8")
    return tmp_path / out


def compute(tmp_path, workbook, out="computed.xlsx"):
    log = tmp_path / "compute.log"
    code = X.main(["--workbook", str(workbook), "--out", str(tmp_path / out), "--log", str(log)], transport=NoJira, command="compute")
    return code, tmp_path / out, log.read_text(encoding="utf-8")


def rebuilt(grids: dict) -> list[Sheet]:
    """The sheets read from a file, to write them again (an empty header cell, between two lists of MetricItems, is the empty text)."""
    return [Sheet(name, tuple("" if c is None else c for c in rows[0]), [tuple(r) for r in rows[1:]]) for name, rows in grids.items() if rows]


def same_sheets(a, b) -> list[str]:
    left, right = read_sheets(a), read_sheets(b)
    assert set(left) == set(right)
    return [name for name in sorted(set(left) - LEFT_OUT) if left[name] != right[name]]


@pytest.fixture(scope="module")
def generated():
    return generate(seed=11, issues=60, projects=project_keys())


@pytest.mark.parametrize("scenario", [None, "art-metrics"])
def test_compute_from_the_xlsx_of_extract_gives_the_same_sheets_and_calls_no_jira(tmp_path, generated, scenario):
    """The generated data, and the scenario art-metrics (the history of the PI field, the teams, the sub-tasks of the Program)."""
    chosen = get(scenario) if scenario else None
    with MockJira(chosen.dataset if chosen else generated, credentials=(MOCK_EMAIL, MOCK_TOKEN)) as jira:
        book = M.write_params_book(tmp_path / "params.xlsx", jira.base_url, {"ProxyServer": "DIRECT", "MaxRetries": "1"}, chosen)
        extracted = extract(tmp_path, book)
        requests = len(jira.requests)
        code, computed, log = compute(tmp_path, extracted)
        assert len(jira.requests) == requests
    assert code == 0, log
    assert "RUN compute" in log and "END sheets of the extraction" in log
    assert same_sheets(extracted, computed) == []


def test_a_corrected_sheet_workflows_takes_effect_as_a_fresh_extraction_with_it(tmp_path, generated):
    """The equivalence of step C on the Python path: the xlsx of an extraction whose sheet Workflows is replaced (the Wait statuses declared Active:
    Flow Efficiency becomes not computable), computed again, is the xlsx of an extraction made with that sheet."""
    with MockJira(generated, credentials=(MOCK_EMAIL, MOCK_TOKEN)) as jira:
        values = {"ProxyServer": "DIRECT", "MaxRetries": "1"}
        book = M.write_params_book(tmp_path / "params.xlsx", jira.base_url, values)
        extracted = extract(tmp_path, book)
        sheets = M.params_book(jira.base_url, values)
        workflows = next(s for s in sheets if s.name == S.of("Workflows"))
        no_wait = Sheet(workflows.name, workflows.header, [r[:2] + ("Active",) + r[3:] if r[2] == "Wait" else r for r in workflows.rows])
        assert no_wait.rows != workflows.rows
        fresh_book = write_xlsx(tmp_path / "no_wait.xlsx", [no_wait if s.name == workflows.name else s for s in sheets])
        fresh = extract(tmp_path, fresh_book, "fresh.xlsx")
    grids = read_sheets(extracted)
    grids[workflows.name] = read_sheets(fresh)[workflows.name]
    edited = write_xlsx(tmp_path / "edited.xlsx", rebuilt(grids))
    code, computed, log = compute(tmp_path, edited)
    assert code == 0, log
    assert same_sheets(fresh, computed) == []
    efficiency = read_sheets(computed, {S.of("Metrics")})[S.of("Metrics")]
    column = efficiency[0].index("FlowEfficiency")
    assert {r[column] for r in efficiency[1:]} == {"not computable"}                     # no Wait status: the correction took effect


def test_compute_needs_no_connection_and_the_extraction_still_does(tmp_path, generated):
    with MockJira(generated, credentials=(MOCK_EMAIL, MOCK_TOKEN)) as jira:
        extracted = extract(tmp_path, M.write_params_book(tmp_path / "params.xlsx", jira.base_url, {"ProxyServer": "DIRECT", "MaxRetries": "1"}))
    grids = read_sheets(extracted)
    config = grids[S.of("Config")]
    value = config[0].index("Value")
    for row in config[1:]:
        if row[0] in ("BaseUrl", "Email", "ApiToken"):
            row[value] = None
    edited = write_xlsx(tmp_path / "no_connection.xlsx", rebuilt(grids))
    code, _out, log = compute(tmp_path, edited)
    assert code == 0, log
    code = X.main(["--workbook", str(edited), "--out", str(tmp_path / "x.xlsx"), "--log", str(tmp_path / "x.log")], detect=direct)
    assert code == 1 and "J2D-CFG-002" in (tmp_path / "x.log").read_text(encoding="utf-8")


def test_a_workbook_with_no_extraction_says_to_extract_first(tmp_path):
    book = M.write_params_book(tmp_path / "params.xlsx", "https://example.invalid", {"ProxyServer": "DIRECT"})
    code, out, log = compute(tmp_path, book)
    assert code == 1 and not out.exists()
    assert "J2D-FLOW-001" in log and "run 1. Extract from Jira first" in log


def test_compute_checks_the_mandatory_parameters_of_the_metrics_as_the_vba_button():
    import re
    from pathlib import Path
    vba = (Path(__file__).parents[2] / "vba" / "modActions.bas").read_text(encoding="utf-8")
    assert tuple(re.search(r'COMPUTE_PARAMETERS As String = "([^"]*)"', vba).group(1).split(",")) == X.COMPUTE_PARAMETERS


# --- the sheets read back -----------------------------------------------------------------------------------------------------------

def test_the_cells_of_a_file_come_back_as_the_extraction_wrote_them():
    import datetime as dt
    assert recompute.from_serial(46323.5) == dt.datetime(2026, 10, 28, 12)
    assert recompute.from_serial(46323 + 1 / 86400 * 0.6) == dt.datetime(2026, 10, 28, 0, 0, 1)         # to the second
    assert recompute.from_serial(None) is None and recompute.from_serial("") is None
    assert recompute.as_text(10001.0) == "10001" and recompute.as_text("TTT-1") == "TTT-1" and recompute.as_text("") is None


def test_a_raw_sheet_without_a_column_names_it():
    with pytest.raises(ParamsError) as err:
        recompute.issue_rows([["Key", "Id"], ["A-1", "1"]])
    assert err.value.code == "J2D-FLOW-003" and "CreatedUtc" in str(err.value)


def test_the_status_changes_and_the_pi_changes_come_from_the_changelog_a_date_that_cannot_be_read_is_left_out():
    import datetime as dt
    day = dt.datetime(2026, 10, 1)
    row = ("A-1", "1", "A", "Story", "Feature", day, "Done", None, "done", None, "R1, R2", None, None, None, None, None, None, None, 0, None, None, 0)
    changelog = [("A-1", "1", "11", day, "x", "status", "status", "jira", "1", "To do", "3", "Done"),
                 ("A-1", "1", "12", None, "garbage", "status", "status", "jira", "3", "Done", "1", "To do"),
                 ("A-1", "1", "13", day, "x", "customfield_9", "PI", "custom", None, None, None, "PI-1")]
    flow = recompute.flow_of([row], changelog, [("A-1", "B-2")], [("A", "Team", "Alpha")], "customfield_9", None)[0]
    assert [(c.id, c.from_status, c.to_status) for c in flow.changes] == [("11", "To do", "Done")]
    assert [(p.id, p.from_value, p.to_value) for p in flow.pi_changes] == [("13", None, "PI-1")]
    assert flow.fix_versions == ("R1", "R2") and flow.blocks == ("B-2",) and flow.team == "Alpha"
    assert recompute.unreadable(changelog) == (1, "String:aaaaaaa x 1")


def test_the_module_runs_the_compute_command(monkeypatch, capsys):
    import runpy
    import sys
    monkeypatch.setattr(sys, "argv", ["compute", "--help"])
    with pytest.raises(SystemExit) as stop:
        runpy.run_module("jira2dashboard.standalone.compute", run_name="__main__")
    assert stop.value.code == 0 and "run.ps1 compute" in capsys.readouterr().out
