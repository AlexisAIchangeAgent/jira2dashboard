"""The scenario art-metrics (mockjira/art_metrics.py): 25 Features over 4 PIs, with the history AL described.

Every number here is typed from AL's text, never computed with the module under test or with flow/: this is the oracle (L0) that
the reference, the standalone extraction and the workbook are later compared with.
"""
import datetime as dt

import pytest

from jira2dashboard.mockjira import api, art_metrics as scenario_module
from jira2dashboard.mockjira.scenarios import get
from jira2dashboard.params.model import project_keys

# --- AL's calendar: [Start, End), PIs of 18 weeks, iterations of 3 weeks, no gap -------------------------------------------------
PIS = {"PI-1": ("2025-01-01", "2025-05-07"), "PI-2": ("2025-05-07", "2025-09-10"),
       "PI-3": ("2025-09-10", "2026-01-14"), "PI-4": ("2026-01-14", "2026-05-20")}
ITERATION_STARTS = {
    "IT-1-1": "2025-01-01", "IT-1-2": "2025-01-22", "IT-1-3": "2025-02-12", "IT-1-4": "2025-03-05", "IT-1-5": "2025-03-26", "IT-1-6": "2025-04-16",
    "IT-2-1": "2025-05-07", "IT-2-2": "2025-05-28", "IT-2-3": "2025-06-18", "IT-2-4": "2025-07-09", "IT-2-5": "2025-07-30", "IT-2-6": "2025-08-20",
    "IT-3-1": "2025-09-10", "IT-3-2": "2025-10-01", "IT-3-3": "2025-10-22", "IT-3-4": "2025-11-12", "IT-3-5": "2025-12-03", "IT-3-6": "2025-12-24",
    "IT-4-1": "2026-01-14", "IT-4-2": "2026-02-04", "IT-4-3": "2026-02-25", "IT-4-4": "2026-03-18", "IT-4-5": "2026-04-08", "IT-4-6": "2026-04-29"}
AS_OF = dt.datetime(2026, 5, 20)

# --- AL's history: where each Feature is released or deferred, and who owns it -----------------------------------------------------
RELEASED = {1: "IT-1-1", 2: "IT-1-2", 3: "IT-1-3", 4: "IT-1-4", 5: "IT-1-5", 6: "IT-1-6",
            7: "IT-2-1", 8: "IT-2-2", 9: "IT-2-3", 10: "IT-2-4", 11: "IT-2-5",
            16: "IT-3-4", 17: "IT-3-4", 18: "IT-3-6",
            19: "IT-4-1", 20: "IT-4-2", 21: "IT-4-3", 22: "IT-4-5", 25: "IT-4-6"}
DEFERRED = {12: "IT-2-6", 13: "IT-3-1", 14: "IT-3-2", 15: "IT-3-3", 23: "IT-4-4", 24: "IT-4-6"}
PI_AFTER_DEFERRAL = {12: None, 13: "PI-5", 14: None, 15: "PI-5", 23: None, 24: "PI-5"}
ENGAGED_AT_START = {"PI-1": range(1, 7), "PI-2": range(7, 13), "PI-3": range(13, 19), "PI-4": range(19, 25)}      # ft-25 comes later
TEAM = {1: "team-1", 2: "team-1", 3: "team-2", 4: "team-2", 5: "team-3", 6: "team-3",
        7: "team-1", 8: "team-2", 9: "team-3", 10: "team-1", 11: "team-2", 12: "team-3",
        13: "team-1", 14: "team-2", 15: "team-3", 16: "team-1", 17: "team-2", 18: "team-3",
        19: "team-1", 20: "team-2", 21: "team-3", 22: "team-1", 23: "team-2", 24: "team-3", 25: "team-1"}
LONG_STATUS = {"PI-1": "testing", "PI-2": "analyzing", "PI-3": "doing", "PI-4": "ready2deploy"}      # 2 days, every other status 1 day
CHAIN = ["analyzing", "ready4PI", "doing", "ready2test", "testing", "ready2deploy", "deploying"]


def day(text: str) -> dt.datetime:
    return dt.datetime.fromisoformat(text)


def iteration_of(moment: dt.datetime) -> str:
    starts = sorted((day(start), label) for label, start in ITERATION_STARTS.items())
    return [label for start, label in starts if start <= moment][-1]


def pi_of_feature(number: int) -> str:
    return next(pi for pi, numbers in ENGAGED_AT_START.items() if number in numbers) if number != 25 else "PI-4"


@pytest.fixture(scope="module")
def scenario():
    return get("art-metrics")


@pytest.fixture(scope="module")
def features(scenario):
    return {i.number: i for i in scenario.dataset.issues}


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


def test_the_fake_jira_holds_only_the_25_features_of_one_project(scenario):
    issues = scenario.dataset.issues
    assert [i.number for i in issues] == list(range(1, 26))
    assert {i.project for i in issues} == {project_keys()[0]}                                    # the Program project, nothing else
    assert {(i.type, i.work_type, i.workflow) for i in issues} == {("Feature", "Feature", "Feature")}
    assert [i.summary for i in issues] == [f"ft-{n}" for n in range(1, 26)]
    assert [i.key for i in issues] == [f"{project_keys()[0]}-{n}" for n in range(1, 26)]


def test_the_workflow_is_program_workflow_with_its_categories(scenario):
    ds = scenario.dataset
    expected = {"todo": "Backlog", "analyzing": "Active", "ready4PI": "Wait", "doing": "Active", "ready2test": "Wait",
                "testing": "Active", "ready2deploy": "Wait", "deploying": "Active", "released": "Done", "deferred": "Backlog"}
    assert {status: ds.catalog.categories[("Feature", status)] for status in expected} == expected
    assert ds.catalog.end_statuses["Feature"] == "released"
    assert ds.workflow_names == {"Feature": "program-workflow"}
    workflows = next(s for s in ds.sheets if s.name == "Workflows")
    feature = [row for row in workflows.rows if row[0] == "Feature"]
    assert [(row[2], row[4], row[5], row[6]) for row in feature] == [          # status, IsStart, IsEnd, IsFree
        ("todo", 0, 0, 0), ("analyzing", 1, 0, 0), ("ready4PI", 0, 0, 0), ("doing", 0, 0, 0), ("ready2test", 0, 0, 0),
        ("testing", 0, 0, 0), ("ready2deploy", 0, 0, 0), ("deploying", 0, 0, 0), ("released", 0, 1, 0), ("deferred", 0, 0, 1)]
    assert any(row[0] == "Team" for row in workflows.rows)                       # the workbook still needs the Team workflow


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
    long_status = LONG_STATUS[pi_of_feature(number)]
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
        assert features[number].pi == pi_of_feature(number)


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
    [feature] = config.project_statuses(ds, program)
    assert [s["name"] for s in feature["statuses"]][-1] == "deferred"


def test_the_mock_jira_command_and_the_workbook_fill_know_the_scenario():
    from jira2dashboard import mock_workbook as M
    from jira2dashboard.mockjira.__main__ import dataset_for, parse
    assert len(dataset_for(parse(["--scenario", "art-metrics"])).issues) == 25
    scenario_ = get("art-metrics")
    book = {sheet.name.split("_")[0]: sheet for sheet in M.params_book("http://x", scenario=scenario_)}
    assert [row[0] for row in book["PIs"].rows] == list(PIS)
    assert len(book["Iterations"].rows) == 24 and book["Iterations"].rows[0][0] == "IT-1-1"
    assert [row[2] for row in book["Workflows"].rows if row[0] == "Feature"][-1] == "deferred"
    config = {row[0]: row[1] for row in book["Config"].rows}
    assert config["AsOf"] == "2026-05-20 00:00" and config["PIField"] == scenario_.dataset.fields.pi
    assert "program-workflow" not in str(book["Workflows"].rows)               # the sheets keep the workflow value Feature


def test_an_unknown_scenario_lists_the_known_ones():
    with pytest.raises(ValueError, match="art-metrics"):
        get("nothing")
