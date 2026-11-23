"""What Jira says about workflows (AL, 2026-11-18): the depth of a status in a flow, the workflows of the issue types, the reading of the preview."""
import pytest

from jira2dashboard.flow.jira_workflows import JiraWorkflows
from jira2dashboard.flow.status_rules import flow_depths
from jira2dashboard.standalone import jira_api as J
from jira2dashboard.standalone.workflows import parse_preview, read_workflows


# --- flow_depths: the moves from the creation status, by breadth-first search ----------------------------------------------------------

def test_the_depth_is_the_number_of_moves_from_the_creation_status():
    moves = [("todo", "analysis"), ("analysis", "doing"), ("doing", "done"), ("analysis", "done")]
    assert flow_depths("todo", moves) == {"todo": 0, "analysis": 1, "doing": 2, "done": 2}               # done: the short way wins


def test_a_loop_is_visited_once_and_a_status_no_move_reaches_has_no_depth():
    assert flow_depths("a", [("a", "b"), ("b", "a"), ("b", "c"), ("x", "a")]) == {"a": 0, "b": 1, "c": 2}      # x leads to a, nothing leads to x


@pytest.mark.parametrize("initial", [None, ""])
def test_without_a_creation_status_there_is_no_depth(initial):
    assert flow_depths(initial, [("a", "b")]) == {}


# --- JiraWorkflows -----------------------------------------------------------------------------------------------------------------

INFO = JiraWorkflows(initial={"WF-A": "Open", "WF-B": "New"}, moves=frozenset({("WF-A", "Open", "Doing"), ("WF-A", "Doing", "Closed"),
                                                                              ("WF-B", "New", "Review"), ("WF-B", "Review", "Doing")}),
                     workflow_of={("PRG", "Feature"): "WF-A", ("PRG", "Epic"): "WF-B", ("TEAM", "Story"): "WF-B"})


def test_a_workflow_is_known_once_its_moves_were_read_and_a_move_is_allowed_inside_it_only():
    assert INFO.knows("WF-A") and not INFO.knows("WF-C")
    assert INFO.allowed("WF-A", "Open", "Doing") and not INFO.allowed("WF-A", "Open", "Closed") and not INFO.allowed("WF-B", "Open", "Doing")


def test_the_depths_of_the_features_come_from_the_workflows_of_their_types_in_the_program_projects_only():
    assert INFO.feature_depths(["PRG"], ["Feature"]) == {"Open": 0, "Doing": 1, "Closed": 2}
    assert INFO.feature_depths(["PRG"], ["Feature", "Epic"]) == {"Open": 0, "Doing": 1, "Closed": 2, "New": 0, "Review": 1}   # the smallest depth wins
    assert INFO.feature_depths(["PRG"], ["Story"]) == {}                                                   # Story belongs to TEAM, not to the Program
    assert INFO.feature_depths([], ["Feature"]) == {} and INFO.feature_depths(["PRG"], ["Unknown"]) == {}


# --- the reading of the preview --------------------------------------------------------------------------------------------------------

PREVIEW = {"statuses": [{"statusReference": "r1", "name": "Open", "statusCategory": "TODO"}, {"statusReference": "r2", "name": "Doing", "statusCategory": "IN_PROGRESS"},
                        {"statusReference": "r3", "name": "Closed", "statusCategory": "DONE"}, {"statusReference": "r4", "name": "Parked", "statusCategory": "TODO"}],
           "workflows": [{"id": "w1", "name": "WF-A", "queryContext": [{"project": "P", "issueTypes": ["10001", "10009"]}],
                          "statuses": [{"statusReference": r} for r in ("r1", "r2", "r3", "r4")],
                          "transitions": [{"type": "INITIAL", "name": "Create", "toStatusReference": "r1", "links": []},
                                          {"type": "DIRECTED", "name": "Start", "id": "11", "toStatusReference": "r2", "links": [{"fromStatusReference": "r1"}]},
                                          {"type": "DIRECTED", "name": "Close", "id": "12", "toStatusReference": "r3", "links": [{"fromStatusReference": "r2"}]},
                                          {"type": "GLOBAL", "name": "Park", "toStatusReference": "r4", "links": []}]}]}


def read(answer, seen=None):
    out = ({}, seen if seen is not None else set(), {}, set(), set(), set())
    parse_preview(answer, "PRG", {"10001": "Feature", "10002": "Bug"}, *out)
    return out


def test_a_preview_gives_the_workflow_of_the_types_its_creation_status_its_moves_its_start_and_its_free_statuses():
    workflow_of, _seen, initial, moves, starts, free = read(PREVIEW)
    assert workflow_of == {("PRG", "Feature"): "WF-A"}                                    # 10009 is not a type of the project: ignored
    assert initial == {"WF-A": "Open"} and starts == {"Doing"} and free == {"Parked"}
    assert moves == {("WF-A", "Open", "Doing"), ("WF-A", "Doing", "Closed")}


def test_a_workflow_shared_by_two_projects_is_read_once_but_each_project_gets_its_types():
    seen = set()
    first = read(PREVIEW, seen)
    second = ({}, seen, {}, set(), set(), set())
    parse_preview(PREVIEW, "OTHER", {"10001": "Task"}, *second)
    assert second[0] == {("OTHER", "Task"): "WF-A"} and second[2] == {} and second[3] == set() and first[3]


class Client:
    def __init__(self, answers):
        self.answers, self.calls = list(answers), []

    def call(self, method, path, what, body=None):
        self.calls.append((method, path.split("?")[0]))
        answer = self.answers.pop(0)
        if isinstance(answer, Exception):
            raise answer
        return answer


PROJECTS = {"isLast": True, "values": [{"id": "10100", "key": "PRG", "issueTypes": [{"id": "10001", "name": "Feature"}]}]}


def test_the_workflows_of_the_projects_are_read_with_the_projects_call_then_the_preview():
    client = Client([PROJECTS, PREVIEW])
    info = read_workflows(client, ["PRG"])
    assert client.calls == [("GET", "/rest/api/3/project/search"), ("POST", "/rest/api/3/workflows/preview")]
    assert info.workflow_of == {("PRG", "Feature"): "WF-A"} and info.starts == {"Doing"} and info.free == {"Parked"}
    assert info.feature_depths(["PRG"], ["Feature"]) == {"Open": 0, "Doing": 1, "Closed": 2}


def test_a_projects_page_that_is_not_the_last_asks_for_the_next_one():
    client = Client([{"isLast": False, "values": [{"id": "1", "key": "A", "issueTypes": []}]}, {"isLast": True, "values": []}])
    assert read_workflows(client, ["A"]).workflow_of == {} and len(client.calls) == 2


def test_a_project_whose_workflows_are_refused_does_not_stop_the_next_one_and_the_warning_counts_it(caplog):
    """As the VBA (modJiraConfig.ReadWorkflowRules): the first project refused (no project administrator right), the second one read."""
    two = {"isLast": True, "values": [{"id": "1", "key": "TEAM", "issueTypes": [{"id": "10002", "name": "Story"}]}] + PROJECTS["values"]}
    client = Client([two, J.JiraError(602, "HTTP 403: no right"), PREVIEW])
    info = read_workflows(client, ["TEAM", "PRG"])
    assert info.starts == {"Doing"} and info.workflow_of == {("PRG", "Feature"): "WF-A"}
    assert "J2D-JIRA-030 | the workflows of Jira could not be read (1 project(s) refused)" in caplog.text


def test_another_error_stops_the_reading_and_the_warning_gives_its_support_code(caplog):
    client = Client([PROJECTS, J.JiraError(602, "HTTP 500: Jira is down")])
    assert read_workflows(client, ["PRG"]) == JiraWorkflows()
    assert "J2D-JIRA-030 | the workflows of Jira could not be read (J2D-JIRA-500)" in caplog.text


def test_a_refused_call_gives_nothing_and_never_raises():
    client = Client([PROJECTS, J.JiraError(602, "no right")])
    assert read_workflows(client, ["PRG"]) == JiraWorkflows()


# --- DC28 on the scenario art-metrics: the sheet Transitions is empty, Jira alone knows the moves ------------------------------------------

def dc28_keys(with_jira):
    from jira2dashboard.flow.checks import CheckSettings, Moves, run_checks
    from jira2dashboard.mockjira.flow_adapter import jira_rules, jira_workflows, to_flow_issues
    from jira2dashboard.mockjira.scenarios import get
    ds = get("art-metrics").dataset
    sheets = list(ds.sheets)
    moves = Moves.from_params(sheets, jira_workflows(ds) if with_jira else JiraWorkflows())
    flags = run_checks(to_flow_issues(ds), ds.as_of, jira_rules(ds), moves, CheckSettings.from_params(sheets))
    return sorted(key for key, ids in flags.items() if "DC28" in ids)


def test_the_only_moves_the_jira_workflow_does_not_have_are_those_of_the_three_subtasks_that_skip_steps():
    assert dc28_keys(with_jira=True) == ["TTT-26", "TTT-27", "TTT-28"]          # analyzing -> doing and doing -> released: no such move; the 25 Features follow Jira


def test_without_the_jira_workflow_and_with_an_empty_sheet_transitions_nothing_can_be_judged():
    assert dc28_keys(with_jira=False) == []
