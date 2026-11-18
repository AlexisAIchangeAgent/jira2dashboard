"""VelocityCheck (flow/velocity_check.py): the Velocity of an iteration counted again from the status changes of the changelog (AL, 2026-11-18).

Why: on the client PC the Velocity of the ART said 60 for an iteration where Jira's `status CHANGED TO "Terminé" DURING` said 8. The table tells
the causes apart: items created already done (no change in the changelog), items closed then reopened, other types than the Features.
Every figure is typed by hand from the small Features below (iterations I1 = 05-01 .. 05-11, I2 = 05-11 .. 05-21, End excluded).
"""
import datetime as dt

from jira2dashboard.flow.itemflow import item_flow
from jira2dashboard.flow.model import OTHER_TYPES, Change, FlowIssue, Iteration, Rules
from jira2dashboard.flow.velocity_check import VelocityCheck, velocity_check
from jira2dashboard.mockjira.scenarios import get

I1 = Iteration("I1", dt.datetime(2025, 5, 1), dt.datetime(2025, 5, 11))
I2 = Iteration("I2", dt.datetime(2025, 5, 11), dt.datetime(2025, 5, 21))
FEATURE_TYPES = frozenset({"Feature"})
AS_OF = dt.datetime(2025, 12, 1)


def day(text):
    return dt.datetime.fromisoformat(text)


def issue(key, moves, initial="todo", created="2025-04-01T00:00", type_="Feature", team="ART"):
    """A Feature of the workflow of the scenario art-metrics (released is its IsEnd status), created in `initial`, that enters each (time, status)."""
    changes, previous = [], initial
    for n, (at, status) in enumerate(moves, start=1):
        changes.append(Change(f"{key}-{n}", day(at), previous, status))
        previous = status
    return FlowIssue(key=key, team=team, type=type_, work_type="Feature", workflow="Feature", created=day(created),
                     status=previous, changes=tuple(changes))


ISSUES = [
    issue("F-1", [("2025-05-02T00:00", "analyzing"), ("2025-05-05T00:00", "released")]),                       # closed in I1
    issue("F-2", [], initial="released", created="2025-05-03T00:00"),                                          # created already done, in I1
    issue("F-3", [("2025-05-04T00:00", "released"), ("2025-05-06T00:00", "doing"), ("2025-05-15T00:00", "released")]),   # closed in I1, reopened, closed in I2
    issue("F-4", [("2025-05-11T00:00", "released")]),                                                          # exactly at the End of I1: I2
    issue("F-5", [("2025-05-02T00:00", "released"), ("2025-05-03T00:00", "doing"), ("2025-05-04T00:00", "released")]),   # closed twice in I1
    issue("E-6", [("2025-05-06T00:00", "released")], type_="Epic", team=OTHER_TYPES),                          # a Program ticket that is not a Feature
    issue("F-7", [("2025-05-07T00:00", "released")], team="team-2"),                                           # another team
    issue("E-9", [("2025-05-08T00:00", "released")], type_="Epic", team="team-2"),                             # an Epic of a Team project: counted, not a Feature
    issue("F-8", [("2025-05-05T00:00", "doing")], initial="released", created="2025-05-03T00:00"),            # created done, reopened at once: not done, not counted
]


def figures(iteration, team="ART"):
    sheets = get("art-metrics").dataset.sheets
    rules = Rules.from_params(list(sheets))
    items = [item_flow(i, AS_OF, rules) for i in ISSUES]
    return velocity_check(ISSUES, items, iteration, team, rules, FEATURE_TYPES)


def test_the_first_iteration_tells_the_creations_the_reopenings_and_the_other_types_apart():
    got = figures(I1)
    # Velocity: F-1, F-2, F-5 (E-6 is of the team "ART (other types)": left out, counted in OtherTypes). Created done: F-2. Changes into released
    # (Seq > 1): F-1, F-3, F-5 twice = 4; reopened later: F-3's and F-5's first = 2. Gap: 3 - 1 - (4 - 2) = 0.
    assert got == VelocityCheck("I1", "ART", velocity=3, velocity_features=3, other_types=1, created_done=1, changed_to_end=4,
                                changed_to_end_features=4, reopened_later=2, gap=0)


def test_the_second_iteration_counts_the_last_close_of_a_reopened_item_and_a_close_at_the_end_of_the_previous_one():
    got = figures(I2)
    # F-3 (closed again 05-15) and F-4 (05-11, the End of I1 is excluded): 2 changes, both final.
    assert got == VelocityCheck("I2", "ART", velocity=2, velocity_features=2, other_types=0, created_done=0, changed_to_end=2,
                                changed_to_end_features=2, reopened_later=0, gap=0)


def test_a_team_counts_only_its_own_items():
    # F-7 (a Feature) and E-9 (an Epic of the Team project: counted in Velocity, not in VelocityFeatures); OtherTypes is on the ART rows only
    assert figures(I1, "team-2") == VelocityCheck("I1", "team-2", velocity=2, velocity_features=1, other_types=0, created_done=0, changed_to_end=2,
                                                  changed_to_end_features=1, reopened_later=0, gap=0)
    assert figures(I2, "team-2") == VelocityCheck("I2", "team-2", 0, 0, 0, 0, 0, 0, 0, 0)


def test_an_iteration_without_any_change_is_all_zero_never_not_computable():
    far = Iteration("I9", dt.datetime(2026, 1, 1), dt.datetime(2026, 1, 11))
    assert figures(far) == VelocityCheck("I9", "ART", 0, 0, 0, 0, 0, 0, 0, 0)


def test_the_gap_is_zero_on_the_real_scenario_whatever_the_iteration():
    scenario = get("art-metrics")
    ds = scenario.dataset
    rules = Rules.from_params(list(ds.sheets))
    from jira2dashboard.mockjira.flow_adapter import to_flow_issues
    issues = to_flow_issues(ds)
    items = [item_flow(i, ds.as_of, rules) for i in issues]
    rows = [velocity_check(issues, items, it, "ART", rules, FEATURE_TYPES) for it in scenario.iterations]
    assert all(r.gap == 0 for r in rows)
    assert all(r.changed_to_end == r.velocity and r.created_done == 0 and r.reopened_later == 0 for r in rows)    # no creation done, none reopened
