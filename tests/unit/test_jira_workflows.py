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


def test_the_keys_jira_did_not_give_back_are_named_in_the_log_and_sent_in_capitals(caplog):
    """AL, 2026-11-26: the example keys left in Projects_from_user gave Projects 0, Workflows 0 and a count only. The keys go in capitals
    (Jira Cloud's keys are; JQL ignores the case, project/search is not said to) and the ones Jira did not give back are named, case ignored."""
    client = Client([PROJECTS, PREVIEW])
    calls = []
    client.call = lambda method, path, what, body=None, _call=client.call: calls.append(path) or _call(method, path, what, body)
    read_workflows(client, ["prg", "TTT", "Phoenix train"])
    assert calls[0].startswith("/rest/api/3/project/search?keys=PRG&keys=TTT&keys=PHOENIX%20TRAIN&")
    assert ("J2D-CFG-016 | project keys not found in Jira: 2 of 3 (sheet Projects_from_user: the keys, not the names): TTT, "
            "Phoenix train") in caplog.text


def test_the_keys_not_found_are_warned_even_when_a_preview_fails_afterwards(caplog):
    """As the VBA (ReadWorkflowRules): the keys are known once project/search has answered, so a preview that fails later loses nothing of it."""
    client = Client([PROJECTS, J.JiraError(602, "HTTP 500: Jira is down")])
    assert read_workflows(client, ["PRG", "XX"]) == JiraWorkflows()
    assert "J2D-CFG-016 | project keys not found in Jira: 1 of 2" in caplog.text and ": XX" in caplog.text
    assert "J2D-JIRA-030 | the workflows of Jira could not be read (J2D-JIRA-500)" in caplog.text


def test_missing_keys_keeps_the_order_of_the_sheet_and_ignores_case():
    from jira2dashboard.standalone.workflows import missing_keys
    assert missing_keys(["B", "a", "C"], {"A"}) == ["B", "C"]
    assert missing_keys(["A"], {"a"}) == [] and missing_keys([], set()) == []


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



# --- the sheets: the layout of the VBA, written and read back ------------------------------------------------------------------------------

def test_the_headers_of_the_jira_sheets_are_those_the_vba_writes():
    import re
    from pathlib import Path
    from jira2dashboard.flow import jira_workflows as JW
    from jira2dashboard.flow.status_rules import STATUS_RULES_HEADER
    vba = (Path(__file__).parents[2] / "vba" / "modJiraConfig.bas").read_text(encoding="utf-8")
    extract = (Path(__file__).parents[2] / "vba" / "modExtract.bas").read_text(encoding="utf-8")
    def header(source, sheet):
        found = re.search(rf'"{sheet}", Array\(([^)]*)\)', source)
        return tuple(re.findall(r'"([^"]*)"', found.group(1)))
    assert header(vba, "JiraStatusCatalog") == JW.CATALOG_HEADER
    assert header(vba, "JiraIssueTypes") == JW.ISSUE_TYPES_HEADER
    assert header(vba, "JiraWorkflows") == JW.WORKFLOWS_HEADER
    assert header(vba, "JiraTransitions") == JW.TRANSITIONS_HEADER
    found = re.findall(r'"StatusRules"\)+, Array\(([^)]*)\)', extract)
    assert len(found) == 2 and all(tuple(re.findall(r'"([^"]*)"', f)) == STATUS_RULES_HEADER for f in found)      # both writers of the VBA


def test_the_workflows_written_in_the_sheets_read_back_the_same():
    from jira2dashboard.flow import jira_workflows as JW
    info = JiraWorkflows(initial={"WF-A": "Open", "WF-B": "New"},
                         moves=frozenset({("WF-A", "Open", "Doing"), ("WF-A", "Doing", "Closed"), ("WF-B", "New", "Review"), ("WF-B", "Review", "Doing")}),
                         starts=frozenset({"Doing", "Review"}), free=frozenset({"Closed", "Cancelled"}),       # Cancelled: reached by a global move only
                         workflow_of={("PRG", "Feature"): "WF-A", ("TEAM", "Story"): "WF-B"})
    rows = JW.sheet_rows(info)
    assert rows["JiraWorkflows"][0] == ("WF-A", "Open", "", "", "", 0, "", 1, "")                    # the creation status first, Initial 1
    info = JW.from_sheets(rows["JiraIssueTypes"], rows["JiraWorkflows"], rows["JiraTransitions"])          # now with the statuses of each workflow
    assert ("", "Cancelled", "", "", "", 1, "", "", "") in rows["JiraWorkflows"]                     # a free status no workflow lists
    assert all(len(r) == len(JW.WORKFLOWS_HEADER) for r in rows["JiraWorkflows"])
    assert JW.from_sheets(rows["JiraIssueTypes"], rows["JiraWorkflows"], rows["JiraTransitions"]) == info
    assert JW.catalog_of(JW.catalog_rows({"Open": "new", "Doing": "indeterminate"})) == {"Open": "new", "Doing": "indeterminate"}


def test_the_sheets_written_by_the_vba_are_read_as_the_vba_reads_them():
    """A row of the VBA: IsStart 1 or empty, IsFree 1 or 0, Initial 1 or empty, numbers read back from a file as floats; an empty table keeps one
    blank row."""
    from jira2dashboard.flow import jira_workflows as JW
    workflows = [("Flow", "To do", "Backlog", None, None, 0.0, "new", 1.0, "PRG"), ("Flow", "Doing", "Active", 1.0, None, 0.0, "indeterminate", None, "PRG"),
                 ("Flow", "Done", "Done", None, 1.0, 1.0, "done", None, "PRG"), (None, None, None, None, None, None, None, None, None)]
    info = JW.from_sheets([("PRG", "Feature", "Flow", "no", 1.0, "company-managed"), (None, None, None, None, None, None)], workflows,
                          [("Flow", "To do", "Doing", "Start", "11", "DIRECTED", "PRG")])
    assert info == JiraWorkflows({"Flow": "To do"}, frozenset({("Flow", "To do", "Doing")}), frozenset({"Doing"}), frozenset({"Done"}),
                                 {("PRG", "Feature"): "Flow"}, {"Flow": ("To do", "Doing", "Done")})



def test_the_jira_workflows_of_a_rule_are_those_of_its_kind_that_hold_its_status():
    """The column JiraWorkflows of StatusRules: the same case as T_JiraWorkflowNames of modTests (the VBA twin)."""
    from jira2dashboard.flow import jira_workflows as JW
    info = JiraWorkflows(workflow_of={("PRG", "Feature"): "WF-P", ("TEAM", "Story"): "WF-T", ("TEAM2", "Story"): "WF-T2"},
                         statuses={"WF-P": ("To do", "Doing"), "WF-T": ("To do",), "WF-T2": ("To do", "Doing"), "WF-X": ("Lost",)})
    assert JW.jira_workflow_names(info, "PRG") == {("Feature", "To do"): "WF-P", ("Feature", "Doing"): "WF-P", ("Team", "To do"): "WF-T, WF-T2",
                                                   ("Team", "Doing"): "WF-T2"}
    assert JW.jira_workflow_names(JiraWorkflows(), "PRG") == {}
