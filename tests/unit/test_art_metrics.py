"""The scenario art-metrics (mockjira/art_metrics.py): 25 Features over 4 PIs, with the history AL described.

Every number here is typed from AL's text, never computed with the module under test or with flow/: this is the oracle (L0) that
the reference, the standalone extraction and the workbook are later compared with.
"""
import datetime as dt

import pytest

from jira2dashboard.mockjira import api, art_metrics as scenario_module
from jira2dashboard.mockjira.scenarios import get
from jira2dashboard.params.model import project_keys

from art_oracle import (AS_OF, CHAIN, DEFERRED, ENGAGED_AT_START, ITERATION_STARTS, LONG_STATUS, PI_AFTER_DEFERRAL, PI_OF_FEATURE,
                        PIS, RELEASED, TEAM)


def day(text: str) -> dt.datetime:
    return dt.datetime.fromisoformat(text)


def iteration_of(moment: dt.datetime) -> str:
    starts = sorted((day(start), label) for label, start in ITERATION_STARTS.items())
    return [label for start, label in starts if start <= moment][-1]


@pytest.fixture(scope="module")
def scenario():
    return get("art-metrics")


@pytest.fixture(scope="module")
def features(scenario):
    return {i.number: i for i in scenario.dataset.issues if i.type == "Feature"}


def at_status(issue, status):
    """When the Feature entered `status` (the first time)."""
    return next(c.at for c in issue.changes if c.to_status == status)


def pi_at(issue, moment):
    """The value of the PI field at `moment`, from the history only: the `from` of the first change after it, else the current value."""
    changes = sorted((c for c in issue.field_changes if c.field_id == scenario_module.FIELDS.pi), key=lambda c: c.at)
    later = [c for c in changes if c.at > moment]
    return later[0].from_string if later else issue.pi


def test_the_calendar_has_no_gap_and_its_dates_are_the_ones_of_the_text(scenario):
    assert {p.name: (p.start, p.end) for p in scenario.pis} == {name: (day(a), day(b)) for name, (a, b) in PIS.items()}
    assert [p.name for p in scenario.pis] == list(PIS)
    assert {i.name: i.start for i in scenario.iterations} == {label: day(start) for label, start in ITERATION_STARTS.items()}
    assert [i.name for i in scenario.iterations] == list(ITERATION_STARTS)                       # PI order, then 1..6
    for earlier, later in zip(scenario.iterations, scenario.iterations[1:]):
        assert earlier.end == later.start and earlier.end - earlier.start == dt.timedelta(weeks=3)
    assert scenario.iterations[-1].end == day(PIS["PI-4"][1])
    assert scenario.dataset.as_of == AS_OF


def test_the_fake_jira_holds_the_25_features_and_3_sub_tasks_of_one_project(scenario):
    issues = scenario.dataset.issues
    assert [i.number for i in issues] == list(range(1, 29))
    assert {i.project for i in issues} == {project_keys()[0]}                                    # the Program project, nothing else
    features = [i for i in issues if i.type == "Feature"]
    assert [i.number for i in features] == list(range(1, 26))
    assert {(i.type, i.work_type, i.workflow) for i in features} == {("Feature", "Feature", "Feature")}
    assert [i.summary for i in features] == [f"ft-{n}" for n in range(1, 26)]
    assert [i.key for i in issues] == [f"{project_keys()[0]}-{n}" for n in range(1, 29)]
    subtasks = [i for i in issues if i.type == "Sub-task"]                                      # not Features: the ART leaves them out
    assert [(i.number, i.parent) for i in subtasks] == [(26, "TTT-7"), (27, "TTT-8"), (28, "TTT-16")]
    assert [i.changes[-1].to_status for i in subtasks] == ["released"] * 3
    assert [i.changes[-1].at.isoformat() for i in subtasks] == ["2025-05-17T12:00:00", "2025-05-17T12:00:00", "2025-11-22T12:00:00"]    # IT-2-1, IT-2-1, IT-3-4: 10 d 12 h in


def test_the_workflow_is_program_workflow_with_its_categories(scenario):
    ds = scenario.dataset
    expected = {"todo": "Backlog", "analyzing": "Active", "ready4PI": "Wait", "doing": "Active", "ready2test": "Wait",
                "testing": "Active", "ready2deploy": "Wait", "deploying": "Active", "released": "Done", "deferred": "Backlog"}
    assert {status: ds.catalog.categories[("Feature", status)] for status in expected} == expected
    assert ds.catalog.end_statuses["Feature"] == "released"
    assert ds.workflow_names == {"Feature": "program-workflow"}
    workflows = next(s for s in ds.jira_sheets if s.name == "Workflows")            # what JIRA defines, in full
    feature = [row for row in workflows.rows if row[0] == "Feature"]
    assert [(row[1], row[3], row[4], row[5]) for row in feature] == [          # status, IsStart, IsEnd, IsFree
        ("todo", 0, 0, 0), ("analyzing", 1, 0, 0), ("ready4PI", 0, 0, 0), ("doing", 0, 0, 0), ("ready2test", 0, 0, 0),
        ("testing", 0, 0, 0), ("ready2deploy", 0, 0, 0), ("deploying", 0, 0, 0), ("released", 0, 1, 0), ("deferred", 0, 0, 1)]
    assert not any(row[0] == "Team" for row in workflows.rows)                  # no Team workflow: the scenario has no Team project
    exceptions = next(s for s in ds.sheets if s.name == "Workflows")            # the sheet holds what Jira cannot say: the Wait statuses (AL, 2026-11-18)
    assert exceptions.rows == (("Feature", "ready4PI", "Wait", None, None, None), ("Feature", "ready2test", "Wait", None, None, None),
                               ("Feature", "ready2deploy", "Wait", None, None, None))


@pytest.mark.parametrize("number, label", sorted(RELEASED.items()))
def test_each_released_feature_is_released_during_its_iteration(features, number, label):
    issue = features[number]
    assert issue.status == "released"
    assert iteration_of(at_status(issue, "released")) == label
    assert issue.iteration == ("IT-4-4" if number == 25 else label)            # ft-25 was put in IT-4-4 and released two iterations later


@pytest.mark.parametrize("number, label", sorted(DEFERRED.items()))
def test_each_deferred_feature_is_deferred_during_its_iteration_and_keeps_its_iteration_field(features, number, label):
    issue = features[number]
    assert issue.status == "deferred"
    assert iteration_of(at_status(issue, "deferred")) == label
    assert issue.iteration == label                                              # the Iteration field stays
    assert not any(c.to_status == "released" for c in issue.changes)


@pytest.mark.parametrize("number", range(1, 26))
def test_every_status_lasts_one_day_but_the_long_one_of_the_pi(features, number):
    issue = features[number]
    long_status = LONG_STATUS[PI_OF_FEATURE[number]]
    assert issue.initial_status == "todo"
    chain = [c for c in issue.changes if c.to_status in CHAIN]
    assert [c.to_status for c in chain] == CHAIN[:len(chain)] and len(chain) in (3, 7)              # released: all 7, deferred: up to doing
    assert [c.from_status for c in issue.changes][0] == "todo"
    for entered, left in zip(issue.changes, issue.changes[1:]):                  # every closed status
        expected = 2 if entered.to_status == long_status else 1
        assert left.at - entered.at == dt.timedelta(days=expected), (number, entered.to_status)


def test_a_released_feature_has_a_flow_time_of_eight_days(features):
    for number in RELEASED:
        issue = features[number]
        assert at_status(issue, "released") - at_status(issue, "analyzing") == dt.timedelta(days=8)


def test_the_team_of_each_feature(features):
    assert {n: i.team for n, i in features.items()} == TEAM
    for pi, numbers in ENGAGED_AT_START.items():
        assert {TEAM[n] for n in numbers} == {"team-1", "team-2", "team-3"}, pi                  # every team has a Feature in each PI


def test_the_pi_field_at_the_start_of_each_pi_is_the_commitment(features):
    for pi, (start, _end) in PIS.items():
        engaged = [n for n, issue in features.items() if issue.created <= day(start) and pi_at(issue, day(start)) == pi]
        assert engaged == list(ENGAGED_AT_START[pi]), pi


def test_the_feature_added_during_pi_4_is_created_after_its_start_with_the_pi_and_the_iteration(features):
    issue = features[25]
    assert issue.created == day("2026-02-10T09:00") and iteration_of(issue.created) == "IT-4-2"
    assert issue.iteration == "IT-4-4" and issue.pi == "PI-4" and issue.field_changes == ()          # initial values, no change item
    assert at_status(issue, "released") > day(PIS["PI-4"][0])


def test_the_pi_field_after_a_deferral_and_the_current_value_is_the_last_change(features):
    for number, after in PI_AFTER_DEFERRAL.items():
        issue = features[number]
        assert issue.pi == after
        last = max((c for c in issue.field_changes), key=lambda c: c.at)
        assert last.to_string == after and last.at == at_status(issue, "deferred")           # same instant as the move to deferred
    for number in set(RELEASED) - {25}:                                           # a released Feature keeps its PI
        assert features[number].pi == PI_OF_FEATURE[number]


def test_the_pi_is_set_before_its_pi_starts(features):
    for pi, start in (("PI-2", "2025-05-07"), ("PI-3", "2025-09-10"), ("PI-4", "2026-01-14")):
        for number in ENGAGED_AT_START[pi]:
            first = min(features[number].field_changes, key=lambda c: c.at)
            assert (first.from_string, first.to_string) == (None, pi)
            assert first.at == day(start) - dt.timedelta(days=1) + dt.timedelta(hours=9)             # last day of the previous PI
    assert all(features[n].field_changes == () for n in (1, 2, 3, 4, 5, 6))                      # PI-1: PI set at creation


def test_the_iteration_field_is_in_the_catalog_and_in_the_issues(scenario):
    ds = scenario.dataset
    assert ds.fields.iteration and ds.fields.iteration in [f["id"] for f in ds.fields.catalog()]
    assert api.issue_json(ds.issues[0], ds, [ds.fields.iteration], 0, "http://x")["fields"] == {ds.fields.iteration: "IT-1-1"}


def test_the_pi_changes_are_served_by_the_changelog_like_a_text_field(scenario):
    ds = scenario.dataset
    body = api.bulk_changelog(ds, {"issueIdsOrKeys": [ds.by_key_or_id(f"{project_keys()[0]}-12").id]}, tz=0, bulk_clamp=1000)
    items = [item for log in body["issueChangeLogs"] for history in log["changeHistories"] for item in history["items"]]
    pi_items = [item for item in items if item["fieldId"] == ds.fields.pi]
    assert [(i["fromString"], i["toString"], i["from"], i["to"]) for i in pi_items] == [(None, "PI-2", None, None), ("PI-2", None, None, None)]
    assert {i["field"] for i in pi_items} == {"PI"}
    only = api.bulk_changelog(ds, {"issueIdsOrKeys": [ds.issues[11].id], "fieldIds": [ds.fields.pi]}, tz=0, bulk_clamp=1000)
    assert [i["fieldId"] for log in only["issueChangeLogs"] for h in log["changeHistories"] for i in h["items"]] == [ds.fields.pi] * 2


def test_the_fake_jira_serves_program_workflow_and_its_statuses(scenario):
    from jira2dashboard.mockjira import config
    ds, program = scenario.dataset, project_keys()[0]
    pid = config.project_id(ds, program)
    preview = config.workflows_preview(ds, {"projectId": pid, "issueTypeIds": ["10001"]})
    [workflow] = preview["workflows"]
    assert workflow["name"] == "program-workflow"
    names = {s["statusReference"]: s["name"] for s in preview["statuses"]}
    assert [names[s["statusReference"]] for s in workflow["statuses"]] == [
        "todo", "analyzing", "ready4PI", "doing", "ready2test", "testing", "ready2deploy", "deploying", "released", "deferred"]
    assert any(t["type"] == "GLOBAL" and names[t["toStatusReference"]] == "deferred" for t in workflow["transitions"])    # the free status
    feature, subtask = config.project_statuses(ds, program)                 # the Feature type and the Sub-task type: the same workflow
    assert [s["name"] for s in feature["statuses"]][-1] == "deferred" and subtask["name"] == "Sub-task" and subtask["subtask"]


def test_the_mock_jira_command_and_the_workbook_fill_know_the_scenario():
    from jira2dashboard import mock_workbook as M
    from jira2dashboard.mockjira.__main__ import dataset_for, parse
    assert len(dataset_for(parse(["--scenario", "art-metrics"])).issues) == 25 + 3        # 25 Features and 3 sub-tasks of the Program
    scenario_ = get("art-metrics")
    book = {sheet.name.split("_")[0]: sheet for sheet in M.params_book("http://x", scenario=scenario_)}
    assert [row[0] for row in book["PIs"].rows] == list(PIS)
    assert len(book["Iterations"].rows) == 24 and book["Iterations"].rows[0][0] == "IT-1-1"
    assert [row[1] for row in book["Workflows"].rows if row[0] == "Feature"] == ["ready4PI", "ready2test", "ready2deploy"]       # the exceptions only
    config = {row[0]: row[1] for row in book["Config"].rows}
    assert config["AsOf"] == "2026-05-20 00:00" and config["PIField"] == scenario_.dataset.fields.pi
    assert "program-workflow" not in str(book["Workflows"].rows)               # the sheets keep the workflow value Feature


def test_an_unknown_scenario_lists_the_known_ones():
    with pytest.raises(ValueError, match="art-metrics"):
        get("nothing")


# --- the scenario holds ONLY its own data (AL, 2026-11-17): its own parameter sheets and fields, nothing of params/*.md or the generator ---

def test_the_scenario_has_its_own_projects_workflows_and_calendar_and_nothing_else(scenario):
    from jira2dashboard.params.model import SPECS, shape_problems
    book = {s.name: s for s in scenario.dataset.sheets}
    assert list(book) == [spec.name for spec in SPECS]                                           # every sheet, in the workbook order
    assert book["Projects"].rows == (("TTT", "Program", None),)                                  # no AAA, EEE, ... Team projects
    assert {r[0] for r in book["Workflows"].rows} == {"Feature"} and book["Transitions"].rows == ()          # exceptions only: Jira holds the rest
    jira = {s.name: s for s in scenario.dataset.jira_sheets}
    assert {r[0] for r in jira["Workflows"].rows} == {"Feature"} and {r[0] for r in jira["Transitions"].rows} == {"Feature"}
    assert [(r[0], r[1], r[2]) for r in book["PIs"].rows] == [(n, a[:10], b[:10]) for n, (a, b) in PIS.items()]
    assert [r[0] for r in book["Iterations"].rows] == list(ITERATION_STARTS) and len(book["Iterations"].rows) == 24
    assert shape_problems({name: book[name] for name in ("Projects", "Iterations", "PIs", "Workflows", "Transitions")}) == []


def test_the_scenario_has_no_team_project_so_the_only_team_of_the_flow_tables_is_the_art(scenario):
    projects = next(s for s in scenario.dataset.sheets if s.name == "Projects")
    assert [r for r in projects.rows if r[1] == "Team"] == []
    assert scenario.dataset.projects() == ["TTT"]


def test_the_scenario_fields_are_only_the_ones_it_uses(scenario):
    ds = scenario.dataset
    names = {f["name"] for f in ds.fields.catalog() if f["custom"]}
    assert names == {"Iteration", "Work type", "PI", "Team"}                                      # no story points, value, sprint, targets
    everything = api._all_fields(ds.issues[0], ds, 0)
    assert None not in everything and "customfield_10016" not in everything and "customfield_10020" not in everything
    assert {k for k in everything if k.startswith("customfield")} == {ds.fields.iteration, ds.fields.work_type, ds.fields.pi, ds.fields.team}


def test_the_config_of_the_scenario_leaves_the_fields_it_does_not_use_empty():
    from jira2dashboard import mock_workbook as M
    scenario_ = get("art-metrics")
    values = M.scenario_config(scenario_, "http://x")
    for name in ("StoryPointsField", "BusinessValueField", "SprintField", "TargetStartField", "TargetEndField"):
        assert values[name] in ("", None), name
    assert values["PIField"] == scenario_.dataset.fields.pi and values["TeamField"] == scenario_.dataset.fields.team


def test_the_mock_jira_in_scenario_mode_never_generates_data(monkeypatch):
    from jira2dashboard.mockjira import __main__ as main_module

    def forbidden(*a, **k):
        raise AssertionError("the generator or the anomalies were used")
    monkeypatch.setattr(main_module, "generate", forbidden)
    monkeypatch.setattr(main_module, "with_anomalies", forbidden)
    assert len(main_module.dataset_for(main_module.parse(["--scenario", "art-metrics", "--anomalies", "--issues", "50"])).issues) == 25 + 3


def test_the_mock_jira_of_a_scenario_listens_on_the_port_of_the_scenario_unless_one_is_given():
    from jira2dashboard.mockjira.__main__ import parse, port_for
    assert port_for(parse(["--scenario", "art-metrics"])) == 8101
    assert port_for(parse(["--scenario", "art-metrics", "--port", "8200"])) == 8200
    assert port_for(parse([])) == 8099
