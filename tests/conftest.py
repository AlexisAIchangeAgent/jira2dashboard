"""Shared fixtures: hand-written Jira cases (results known by hand), a tiny HTTP client for the fake Jira, and the
test workbook shared by every `@excel` scenario (built once per test session, opened in one hidden Excel)."""
import datetime as dt
import gc
import json
import urllib.error
import urllib.request

import pytest

from jira2dashboard.excel import com
from jira2dashboard.mockjira.dataset import Dataset, issue
from jira2dashboard.paths import INTER_DIR
from jira2dashboard.testbench import keep_log
from jira2dashboard.testbench import run_vba_tests as rvt

AS_OF = dt.datetime(2026, 11, 3, 18, 0)
SHARED_WORKBOOK = INTER_DIR / "workbook_test.xlsm"
TEST_LOGS = INTER_DIR / "logs"                  # the logs of the test workbooks (all in out\inter), kept by `clean`
REPORTS = pytest.StashKey[dict]()


@pytest.hookimpl(wrapper=True)
def pytest_runtest_makereport(item, call):
    """Keeps the report of each phase on the test: keep_excel_logs writes its result in the header of the kept log."""
    report = yield
    item.stash.setdefault(REPORTS, {})[report.when] = report
    return report


@pytest.fixture(autouse=True)
def keep_excel_logs(request):
    """Each Excel test (marker excel, every module) leaves its logs in out\\inter\\logs\\excel-test_<test>.log: its
    result and time, the lines it added to each log of that folder (the VBA log, vba-tests.log), and the logs of its
    temporary folder (python.log). Read them after a run: build and vba-tests keep that folder (testbench/keep_log.py)."""
    if request.node.get_closest_marker("excel") is None:
        yield
        return
    temporary = request.getfixturevalue("tmp_path")
    marks = keep_log.mark(TEST_LOGS)
    yield
    reports = request.node.stash.get(REPORTS, {})
    report = reports.get("call") or reports.get("setup")
    result = report.outcome if report else "not run"
    took = f"{report.duration:.3f} s" if report else "-"
    header = f"# {request.node.nodeid} | {result} | {took} | {dt.datetime.now():%Y-%m-%d %H:%M:%S}"
    keep_log.keep(TEST_LOGS, marks, request.node.name, header, sorted(temporary.rglob("*.log")))


@pytest.fixture(scope="session")
def built_workbook():
    """The test workbook (VBA modules + parameter sheets), built once and opened in one hidden Excel.

    Yields a holder, not the COM object: pytest keeps the yielded value in its cache until after teardown,
    which would release the COM proxy once Excel has quit (RPC error dumped by faulthandler).
    """
    rvt.build(rvt.ORDER, SHARED_WORKBOOK)
    holder = {}
    with com.excel_session() as app:
        app.AutomationSecurity = com.MSO_AUTOMATION_SECURITY_LOW
        holder["wb"] = app.Workbooks.Open(str(SHARED_WORKBOOK.resolve()))
        yield holder
        holder.pop("wb").Close(False)
        del app                        # release every COM proxy before excel_session quits Excel
        gc.collect()


@pytest.fixture(scope="session")
def small_dataset() -> Dataset:
    """Five issues: a clean run to Terminé, a restart after Dépriorisé, an open one, a cancelled one, a feature."""
    issues = (
        issue("TA-1", id="10001", type="Story", work_type="Feature", story_points=5, assignee="Ann Martin",
              moves=(("2026-10-02T10:00", "En analyse"), ("2026-10-05T10:00", "Prêt"),
                     ("2026-10-06T10:00", "Planifié"), ("2026-10-07T10:00", "En développement"),
                     ("2026-10-12T10:00", "En test"), ("2026-10-13T10:00", "À déployer"),
                     ("2026-10-14T10:00", "Terminé")), parent="PRG-1"),
        issue("TA-2", id="10002", type="Task", moves=(("2026-10-02T10:00", "En analyse"),
                                                         ("2026-10-04T10:00", "Dépriorisé"),
                                                         ("2026-10-06T10:00", "Nouveau"),
                                                         ("2026-10-08T10:00", "En analyse"))),
        issue("TA-3", id="10003", type="Bug", work_type="Defect"),
        issue("TB-1", id="10004", type="Story", moves=(("2026-10-03T10:00", "Dépriorisé"),
                                                         ("2026-10-09T10:00", "Annulé"))),
        issue("PRG-1", id="10005", type="Feature", work_type="Feature", business_value=40, pi="PI#01-2026",
              created="2026-08-15T09:00", moves=(("2026-09-01T10:00", "En analyse"),)),
    )
    return Dataset(issues, AS_OF)


def at(text: str) -> dt.datetime:
    """'10-05' -> 2026-10-05 00:00 UTC."""
    return dt.datetime.fromisoformat(f"2026-{text}")


def flow_issue(key, team, type_, work_type, created, moves, workflow="Team"):
    """A FlowIssue from hand-written moves: ('10-02', 'En analyse') means: entered En analyse on 2026-10-02."""
    from jira2dashboard.flow.model import Change, FlowIssue
    changes, previous = [], "Nouveau"
    for n, (day, status) in enumerate(moves, start=1):
        changes.append(Change(f"{key}-{n}", at(day), previous, status))
        previous = status
    return FlowIssue(key, team, type_, work_type, workflow, at(created), previous, tuple(changes))


FLOW_AS_OF = dt.datetime(2026, 11, 2)


@pytest.fixture(scope="session")
def flow_cases():
    """Nine issues for the flow metrics, with results worked out by hand in tests/unit/test_flow_reference.py.

    Alpha: A-1 clean run to Terminé; A-2 goes through Depriorise then Nouveau (the flow time restarts); A-3 done
    without ever entering a start status; A-4 reopened after Terminé; A-5 still in progress; A-6 cancelled.
    Beta: B-1 clean run; B-2 passes through a status missing from Workflows (Unclassified).
    ART: PRG-1 a Program feature (Feature workflow). As of 2026-11-02.
    """
    return [
        flow_issue("A-1", "Alpha", "Story", "Feature", "10-01", [
            ("10-02", "En analyse"), ("10-04", "Prêt"), ("10-05", "Planifié"), ("10-06", "En développement"),
            ("10-10", "En test"), ("10-12", "À déployer"), ("10-13", "Terminé")]),
        flow_issue("A-2", "Alpha", "Task", "Autre", "10-01", [
            ("10-02", "En analyse"), ("10-06", "Dépriorisé"), ("10-08", "Nouveau"), ("10-09", "En analyse"),
            ("10-11", "Prêt"), ("10-12", "En développement"), ("10-16", "En test"), ("10-17", "À déployer"),
            ("10-18", "Terminé")]),
        flow_issue("A-3", "Alpha", "Bug", "Defect", "10-03", [("10-04", "Planifié"), ("10-08", "Terminé")]),
        flow_issue("A-4", "Alpha", "Story", "Enabler", "10-01", [
            ("10-02", "En analyse"), ("10-05", "En développement"), ("10-09", "Terminé"), ("10-10", "En test"),
            ("10-14", "Terminé")]),
        flow_issue("A-5", "Alpha", "Story", "Feature", "10-10", [
            ("10-12", "En analyse"), ("10-20", "En développement")]),
        flow_issue("A-6", "Alpha", "Task", "Autre", "10-02", [
            ("10-03", "En analyse"), ("10-07", "Dépriorisé"), ("10-09", "Annulé")]),
        flow_issue("B-1", "Beta", "Story", "Feature", "10-10", [
            ("10-11", "En analyse"), ("10-15", "Prêt"), ("10-16", "Planifié"), ("10-20", "En développement"),
            ("10-25", "En test"), ("10-27", "À déployer"), ("10-28", "Terminé")]),
        flow_issue("B-2", "Beta", "Story", None, "10-12", [
            ("10-13", "En analyse"), ("10-15", "Revue"), ("10-18", "En développement"), ("10-22", "Terminé")]),
        flow_issue("PRG-1", "ART", "Feature", "Feature", "10-01", [
            ("10-02", "En analyse"), ("10-03", "Prêt"), ("10-04", "Planifié"), ("10-05", "En développement"),
            ("10-14", "En test / recette"), ("10-20", "À déployer"), ("10-21", "Terminé")], workflow="Feature"),
    ]


class HttpResult:
    def __init__(self, status, headers, raw):
        self.status, self.headers, self.raw = status, headers, raw

    @property
    def json(self):
        return json.loads(self.raw)


@pytest.fixture
def call():
    """call(jira, method, path, body=None, headers=None) -> HttpResult (errors are results, not exceptions)."""
    def _call(jira, method, path, body=None, headers=None):
        data = json.dumps(body).encode("utf-8") if body is not None else None
        request = urllib.request.Request(jira.base_url + path, data=data, method=method)
        request.add_header("Content-Type", "application/json")
        for name, value in (headers or {}).items():
            request.add_header(name, value)
        try:
            with urllib.request.urlopen(request, timeout=10) as response:
                return HttpResult(response.status, dict(response.headers), response.read())
        except urllib.error.HTTPError as err:
            return HttpResult(err.code, dict(err.headers), err.read())
    return _call
