"""Dates, times and timestamps: the cases that an audit of the VBA and the Python code found wrong or fragile (2026-11-13)."""
import datetime as dt

from jira2dashboard.flow.intervals import build_intervals
from jira2dashboard.flow.model import Change, FlowIssue, Iteration, Rules, change_order
from jira2dashboard.standalone import mapping as M
from jira2dashboard.standalone import results as R


def issue_with(changes):
    return FlowIssue(key="TA-1", team="Team A", type="Story", work_type=None, workflow="Team", created=dt.datetime(2026, 10, 1),
                     status="C", changes=tuple(changes))


def test_two_changes_in_the_same_second_keep_the_order_of_jira_ids_as_numbers_not_as_text():
    at = dt.datetime(2026, 10, 2, 10, 0, 0)                       # the milliseconds are dropped: two changes can tie
    changes = [Change("10", at, "B", "C"), Change("9", at, "A", "B")]
    assert [c.id for c in sorted(changes, key=change_order)] == ["9", "10"]          # a text sort gives 10 before 9
    statuses = [i.status for i in build_intervals(issue_with(changes), dt.datetime(2026, 11, 1), Rules.from_params())]
    assert statuses == ["A", "B", "C"]


def test_the_days_of_the_iterations_run_from_the_earliest_start_to_the_latest_end_in_any_order():
    out_of_order = [Iteration("I2", dt.datetime(2026, 10, 19), dt.datetime(2026, 11, 9)),
                    Iteration("I1", dt.datetime(2026, 10, 5), dt.datetime(2026, 10, 19))]
    first, end = R.iteration_span(out_of_order)
    assert (first, end) == (dt.datetime(2026, 10, 5), dt.datetime(2026, 11, 9))
    assert len(R.days_of(first, end)) == 35                       # the sheet DailyLoad of the VBA holds the same 35 days


def test_an_absurd_date_is_not_a_date_instead_of_a_crash():
    assert M.parse_jira_date("9999-12-31T23:59:59-0100") is None         # the offset pushes it past year 9999
    assert M.parse_jira_date("0001-01-01T00:00:00+0100") is None         # or before year 1
    assert M.parse_jira_date("2026-10-01T09:12:33.123+0530") == dt.datetime(2026, 10, 1, 3, 42, 33)
