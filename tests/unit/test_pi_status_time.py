"""PIStatusTime (flow/pi_status_time.py): the cycle time of the Features, status by status, PI by PI.

On the scenario art-metrics the numbers are the hand-worked ones of tests/art_oracle.py (oracle L0). The hard cases that the scenario
cannot show (an item that crosses a PI, a status visited twice, a backlog return, a reopened item, the edges of the window) are small
hand-built Features whose expected figures are typed here.
"""
import datetime as dt

import pytest

from jira2dashboard.flow.itemflow import item_flow
from jira2dashboard.flow.model import NOT_COMPUTABLE, Change, FlowIssue, Rules
from jira2dashboard.flow.pi_status_time import feature_statuses, pi_status_time, released_in
from jira2dashboard.mockjira.flow_adapter import to_flow_issues
from jira2dashboard.mockjira.scenarios import get

from art_oracle import CHAIN, PI_ITEMS, PIS, STATUS_AVERAGE_DAYS, STATUS_SHARE


def day(text):
    return dt.datetime.fromisoformat(text)


@pytest.fixture(scope="module")
def world():
    scenario = get("art-metrics")
    ds = scenario.dataset
    rules = Rules.from_params(list(ds.sheets))
    items = [item_flow(i, ds.as_of, rules) for i in to_flow_issues(ds)]
    windows = {p.name: (p.start, p.end) for p in scenario.pis}
    statuses = feature_statuses(ds.sheets)
    return {"ds": ds, "items": items, "windows": windows, "statuses": statuses, "rules": rules,
            "rows": pi_status_time(items, statuses, list(windows), windows)}


def test_the_stack_is_the_active_and_wait_statuses_of_the_feature_workflow_in_the_order_of_the_sheet(world):
    assert world["statuses"] == CHAIN                              # todo and deferred are Backlog, released is Done


def test_the_scenario_gives_the_items_the_average_days_and_the_shares_worked_out_by_hand(world):
    rows = world["rows"]
    assert [(r.pi, r.status) for r in rows] == [(pi, s) for pi in PIS for s in CHAIN]                  # PI order, then status order
    assert {r.pi: r.items for r in rows} == PI_ITEMS                                                   # not only the shares: Items cancels out
    assert {(r.pi, r.status): r.average_days for r in rows} == {(pi, s): d for pi, ds in STATUS_AVERAGE_DAYS.items() for s, d in ds.items()}
    assert {(r.pi, r.status): r.share for r in rows} == {(pi, s): v for pi, ds in STATUS_SHARE.items() for s, v in ds.items()}
    assert all(abs(sum(r.share for r in rows if r.pi == pi) - 1.0) < 1e-12 for pi in PIS)              # the bar of a PI is 100 %


def test_the_totals_of_the_statuses_are_the_active_and_wait_days_of_the_same_items(world):
    for pi, window in world["windows"].items():
        mine = released_in(world["items"], window)
        assert abs(sum(r.total_days for r in world["rows"] if r.pi == pi) - sum(i.active_days + i.wait_days for i in mine)) < 1e-9, pi


def test_the_average_days_of_the_statuses_add_up_to_the_average_flow_time(world):
    for pi, window in world["windows"].items():
        mine = released_in(world["items"], window)
        assert sum(r.average_days for r in world["rows"] if r.pi == pi) == pytest.approx(sum(i.flow_time_days for i in mine) / len(mine))


# --- the hard cases, with Features of their own ---------------------------------------------------------------------------------

PI_A, PI_B = (day("2025-05-01"), day("2025-05-11")), (day("2025-05-11"), day("2025-05-21"))
WINDOWS = {"A": PI_A, "B": PI_B}


def feature(key, moves, created="2025-04-01T00:00"):
    """A Feature that enters each (time, status) in turn; it starts in todo."""
    changes, previous = [], "todo"
    for n, (at, status) in enumerate(moves, start=1):
        changes.append(Change(f"{key}-{n}", day(at), previous, status))
        previous = status
    return FlowIssue(key=key, team="ART", type="Feature", work_type="Feature", workflow="Feature", created=day(created),
                     status=previous, changes=tuple(changes))


def compute(*features, names=("A", "B")):
    sheets = get("art-metrics").dataset.sheets
    rules = Rules.from_params(list(sheets))
    items = [item_flow(f, day("2025-12-01"), rules) for f in features]
    rows = pi_status_time(items, feature_statuses(sheets), list(names), WINDOWS)
    return {(r.pi, r.status): r for r in rows}, items


def test_a_feature_that_crosses_the_pi_counts_in_the_pi_where_it_was_released_with_all_its_time():
    # analyzing 2025-05-09 to 05-12 (3 days: two of them before the end of A), doing 05-12 to 05-15, released 05-15 (in B)
    rows, _ = compute(feature("F-1", [("2025-05-09T00:00", "analyzing"), ("2025-05-12T00:00", "doing"), ("2025-05-15T00:00", "released")]))
    assert rows[("A", "analyzing")].items == 0 and rows[("A", "analyzing")].average_days == NOT_COMPUTABLE
    assert (rows[("B", "analyzing")].items, rows[("B", "analyzing")].total_days, rows[("B", "doing")].total_days) == (1, 3.0, 3.0)
    assert rows[("B", "analyzing")].share == 0.5 and rows[("B", "ready4PI")].total_days == 0.0


def test_a_status_visited_twice_adds_up_and_an_item_that_never_visited_a_status_counts_zero():
    moves = [("2025-05-02T00:00", "analyzing"), ("2025-05-04T00:00", "ready4PI"), ("2025-05-05T00:00", "analyzing"),
             ("2025-05-06T00:00", "doing"), ("2025-05-07T00:00", "released")]
    other = [("2025-05-03T00:00", "analyzing"), ("2025-05-04T00:00", "doing"), ("2025-05-06T00:00", "released")]   # 1 day, 2 days, no ready4PI
    rows, _ = compute(feature("F-1", moves), feature("F-2", other))
    assert rows[("A", "analyzing")].total_days == 4.0 and rows[("A", "analyzing")].average_days == 2.0          # (3 + 1) / 2: twice in F-1
    assert rows[("A", "doing")].total_days == 3.0 and rows[("A", "doing")].average_days == 1.5                 # (1 + 2) / 2
    assert rows[("A", "ready4PI")].total_days == 1.0 and rows[("A", "ready4PI")].average_days == 0.5          # F-2 never went there: counts 0
    assert rows[("A", "analyzing")].items == 2
    assert sum(r.share for (pi, _s), r in rows.items() if pi == "A") == pytest.approx(1.0)


def test_the_time_before_a_return_to_the_backlog_is_not_cycle_time():
    # analyzing 05-02 to 05-04, back to todo 05-04 to 05-06, analyzing again 05-06 to 05-08, released 05-08: the flow restarts at 05-06
    moves = [("2025-05-02T00:00", "analyzing"), ("2025-05-04T00:00", "todo"), ("2025-05-06T00:00", "analyzing"), ("2025-05-08T00:00", "released")]
    rows, items = compute(feature("F-1", moves))
    assert rows[("A", "analyzing")].total_days == 2.0 and items[0].start == day("2025-05-06")


def test_a_feature_reopened_after_being_done_is_left_out_as_flow_time_leaves_it_out():
    moves = [("2025-05-02T00:00", "analyzing"), ("2025-05-03T00:00", "doing"), ("2025-05-04T00:00", "released"),
             ("2025-05-05T00:00", "doing"), ("2025-05-07T00:00", "released")]
    other = [("2025-05-02T00:00", "analyzing"), ("2025-05-03T00:00", "doing"), ("2025-05-05T00:00", "released")]
    rows, items = compute(feature("F-1", moves), feature("F-2", other))
    assert items[0].reopened == 1 and rows[("A", "doing")].items == 1 and rows[("A", "doing")].total_days == 2.0       # F-2 only


def test_a_feature_done_without_a_valid_start_is_left_out():
    rows, items = compute(feature("F-1", [("2025-05-03T00:00", "released")]))                  # todo -> released: it never started
    assert items[0].skipped_start and rows[("A", "analyzing")].items == 0 and rows[("A", "analyzing")].share == NOT_COMPUTABLE


def test_the_window_is_half_open_a_feature_released_at_the_end_of_a_pi_belongs_to_the_next_one():
    at_end = feature("F-1", [("2025-05-08T00:00", "analyzing"), ("2025-05-09T00:00", "doing"), ("2025-05-11T00:00", "released")])   # = the end of A
    at_start = feature("F-2", [("2025-04-29T00:00", "analyzing"), ("2025-04-30T00:00", "doing"), ("2025-05-01T00:00", "released")])  # = the start of A
    rows, _ = compute(at_end, at_start)
    assert (rows[("A", "doing")].items, rows[("B", "doing")].items) == (1, 1)
    assert (rows[("A", "doing")].total_days, rows[("B", "doing")].total_days) == (1.0, 2.0)


def test_a_pi_without_a_released_feature_or_without_dates_is_not_computable_never_zero():
    rows, _ = compute(feature("F-1", [("2025-05-02T00:00", "doing"), ("2025-05-04T00:00", "released")]), names=("A", "B", "Z"))
    assert rows[("B", "doing")].items == 0 and rows[("B", "doing")].average_days == NOT_COMPUTABLE and rows[("B", "doing")].share == NOT_COMPUTABLE
    assert all(r.items == r.share == NOT_COMPUTABLE for (pi, _s), r in rows.items() if pi == "Z")            # the PI "Z" has no dates


def test_a_feature_that_is_not_done_is_not_counted():
    rows, _ = compute(feature("F-1", [("2025-05-02T00:00", "doing")]))                          # still in doing at the as-of date
    assert rows[("A", "doing")].items == 0
