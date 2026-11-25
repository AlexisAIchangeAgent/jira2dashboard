"""IterationFeatures and Feature Predictability on the scenario art-metrics, against numbers worked out BY HAND (oracle L0).

The tables below were derived from AL's text and the scenario's facts (tests/unit/test_art_metrics.py), feature by feature, never by
running the code under test. Columns: Engaged, Added, Deferred, Released, ReleasedAdded, CarryOver.
"""
import datetime as dt

import pytest

from jira2dashboard.flow.checks import CheckSettings
from jira2dashboard.flow.itemflow import item_flow
from jira2dashboard.flow.iteration_features import baselines, feature_teams, iteration_features
from jira2dashboard.flow.metrics import PIPredictability, feature_predictability
from jira2dashboard.flow.model import NOT_COMPUTABLE, Iteration, Rules
from jira2dashboard.mockjira.flow_adapter import jira_rules, to_flow_issues
from jira2dashboard.mockjira.scenarios import get

from art_oracle import ART_ROWS, CURRENT_PI, DONE_A, DONE_B, NOT_IN_SHEET, PLANNED, TEAM_ROWS


@pytest.fixture(scope="module")
def world():
    scenario = get("art-metrics")
    ds = scenario.dataset
    issues = to_flow_issues(ds)
    rules = jira_rules(ds)
    items = [item_flow(i, ds.as_of, rules) for i in issues]
    settings = CheckSettings(feature_types=frozenset({"Feature"}), pis={p.name: (p.start, p.end) for p in scenario.pis})
    rows = iteration_features(issues, items, list(scenario.iterations), settings.pis, settings.feature_types)
    return {"scenario": scenario, "issues": issues, "items": items, "settings": settings, "rows": rows}


def row_values(r):
    return (r.engaged, r.added, r.deferred, r.released, r.released_added, r.carry_over)


def test_the_art_rows_are_the_ones_worked_out_by_hand(world):
    art = [r for r in world["rows"] if r.team == "ART"]
    assert [r.iteration for r in art] == list(ART_ROWS)
    assert {r.iteration: row_values(r) for r in art} == ART_ROWS
    assert all(r.pi == r.iteration[3:4].join(["PI-", ""]) for r in art)          # IT-<n>-<k> belongs to PI-<n>
    assert all(r.released_planned + r.released_added == r.released for r in art)


def test_each_team_row_is_the_one_worked_out_by_hand(world):
    rows = {(r.team, r.iteration): (r.engaged, r.added, r.deferred, r.released) for r in world["rows"]}
    for team, expected in TEAM_ROWS.items():
        assert {label: rows[(team, label)] for label in expected} == expected, team


def test_the_art_is_the_sum_of_the_teams_and_the_rows_are_in_order(world):
    rows = world["rows"]
    assert feature_teams(world["issues"], frozenset({"Feature"})) == ["ART", "team-1", "team-2", "team-3"]
    assert [(r.iteration, r.team) for r in rows[:5]] == [("IT-1-1", "ART"), ("IT-1-1", "team-1"), ("IT-1-1", "team-2"),
                                                          ("IT-1-1", "team-3"), ("IT-1-2", "ART")]
    for label in ART_ROWS:
        art = next(r for r in rows if r.iteration == label and r.team == "ART")
        mine = [r for r in rows if r.iteration == label and r.team != "ART"]
        for column in ("engaged", "added", "deferred", "released", "released_planned", "released_added", "carry_over"):
            assert getattr(art, column) == sum(getattr(r, column) for r in mine), (label, column)


def test_what_an_iteration_carries_over_is_what_the_next_one_of_the_pi_starts_with(world):
    for team in ("ART", "team-1", "team-2", "team-3"):
        mine = [r for r in world["rows"] if r.team == team]
        for before, after in zip(mine, mine[1:]):
            if before.pi == after.pi:
                assert before.carry_over == after.engaged, (team, before.iteration)


def test_feature_predictability_counts_what_was_engaged_at_the_start_of_each_pi(world):
    got = feature_predictability(world["issues"], world["items"], world["settings"], world["scenario"].iterations)
    assert got[:4] == [PIPredictability(pi, PLANNED[pi], DONE_A[pi], DONE_A[pi] / PLANNED[pi]) for pi in PLANNED]
    assert [round(p.ratio, 4) for p in got[:4]] == [1.0, 0.8333, 0.5, 0.6667]
    assert got[4:] == [PIPredictability(pi, n, None, NOT_COMPUTABLE) for pi, n in NOT_IN_SHEET.items()]       # PI-5: the sheet lacks it
    both = feature_predictability(world["issues"], world["items"], world["settings"], world["scenario"].iterations, "B")
    assert [(p.planned, p.done) for p in both[:4]] == [(PLANNED[pi], DONE_B[pi]) for pi in PLANNED]
    assert [round(p.ratio, 4) for p in both[:4]] == [1.0, 0.8333, 0.5, 0.8333]


def test_reading_the_pi_field_as_it_is_today_would_lose_the_deferred_features_and_count_the_added_one(world):
    """The pitfall: planned from the current PI field is 6, 5, 3, 5 while 6 Features were engaged at the start of each PI."""
    features = [i for i in world["issues"] if i.type == "Feature"]
    assert {pi: sum(1 for i in features if i.pi == pi) for pi in PLANNED} == CURRENT_PI != PLANNED


def test_the_baseline_of_a_pi_is_the_start_of_its_first_iteration(world):
    scenario = world["scenario"]
    assert baselines(list(scenario.iterations), world["settings"].pis) == {p.name: p.start for p in scenario.pis}
    assert baselines([], world["settings"].pis) == {}                    # no iteration: no baseline


def test_a_pi_without_iteration_has_nothing_planned_a_gap_never_zero(world):
    got = feature_predictability(world["issues"], world["items"], world["settings"], [])
    assert len(got) == 5 and all(p.planned == 0 and p.ratio == NOT_COMPUTABLE for p in got[:4])


def test_the_pis_that_hold_an_iteration_come_by_start_then_in_the_natural_order_of_their_names():
    """AL, 2026-11-25: the rows of an iteration follow the order of PIMetrics (by Start, then the names, numbers as numbers), not the sheet."""
    day = dt.datetime
    window = (day(2026, 1, 1), day(2026, 3, 1))
    pis = {"PI-B": window, "PI-10": window, "PI-9": window, "Late": (day(2026, 2, 1), day(2026, 4, 1)),
           "Early": (day(2025, 12, 1), day(2026, 2, 1))}
    rows = iteration_features([], [], [Iteration("I1", day(2026, 1, 5), day(2026, 1, 19))], pis, frozenset({"Feature"}))
    assert [r.pi for r in rows] == ["Early", "PI-9", "PI-10", "PI-B"]
