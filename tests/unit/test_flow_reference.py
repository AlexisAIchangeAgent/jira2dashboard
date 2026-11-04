"""Flow metrics, reference implementation, on nine hand-written issues (see the flow_cases fixture).

Every expected value below was worked out by hand from docs/safe-flow-metrics-and-data-checks.md section 2; the
arithmetic is in the comments. Dates are 2026, UTC midnight. As of 2026-11-02.
"""
import datetime as dt

import pytest

from conftest import FLOW_AS_OF, at
from jira2dashboard.flow.itemflow import item_flow
from jira2dashboard.flow.load import daily_load, in_process_at, snapshot_days
from jira2dashboard.flow.metrics import flow_metrics, percentile_inc, team_period_metrics
from jira2dashboard.flow.model import NOT_COMPUTABLE, Period, Rules

P1 = Period("P1", at("10-05"), at("10-19"))
P2 = Period("P2", at("10-19"), at("11-02"))


@pytest.fixture(scope="module")
def rules():
    return Rules.from_params()


@pytest.fixture(scope="module")
def items(flow_cases, rules):
    return [item_flow(issue, FLOW_AS_OF, rules) for issue in flow_cases]


@pytest.fixture(scope="module")
def row(items):
    by_key = {i.key: i for i in items}
    return by_key.__getitem__


# --- ItemFlow ---------------------------------------------------------------------------------------

def test_a_clean_run(row):
    a = row("A-1")
    assert (a.start, a.first_start, a.done, a.removed) == (at("10-02"), at("10-02"), at("10-13"), None)
    assert (a.flow_time_days, a.lead_time_days) == (11, 12)               # 10-13 - 10-02 ; 10-13 - 10-01
    assert a.active_days == 8                                              # En analyse 2 + En dev 4 + En test 2
    assert a.wait_days == 3                                                # Pret 1 + Planifie 1 + A deployer 1
    assert (a.unclassified_days, a.backlog_returns, a.reopened, a.skipped_start) == (0, 0, 0, False)


def test_the_last_interval_is_open_and_runs_to_the_as_of_date(row):
    last = row("A-1").intervals[-1]
    assert (last.is_open, last.left, last.status, last.category) == (True, None, "Terminé", "Done")
    assert last.duration_days == 20                                        # 10-13 -> 11-02
    assert [i.seq for i in row("A-1").intervals] == list(range(1, 9))
    assert row("A-1").intervals[0].change_id == "A-1-1" and last.change_id is None


def test_a_return_through_backlog_restarts_the_flow_time(row):
    a = row("A-2")
    assert a.first_start == at("10-02")                    # the first entry into En analyse, ever
    assert a.start == at("10-09")                          # the first one after the last Backlog entry (10-08)
    assert a.backlog_returns == 2                          # Deprioritise (10-06) and Nouveau (10-08) after 10-02
    assert (a.done, a.flow_time_days, a.lead_time_days) == (at("10-18"), 9, 17)     # 10-18 - 10-09 ; 10-18 - 10-01
    assert (a.active_days, a.wait_days) == (7, 2)          # from 10-09: analyse 2 + dev 4 + test 1 ; Pret 1 + deploy 1


def test_done_without_ever_starting_is_a_skipped_start(row):
    a = row("A-3")
    assert (a.start, a.first_start, a.done) == (None, None, at("10-08"))
    assert a.skipped_start and a.flow_time_days is None and a.active_days is None
    assert a.lead_time_days == 5                           # 10-08 - 10-03
    assert a.excluded_from_flow_time


def test_reopened_after_terminated_keeps_the_last_done_date(row):
    a = row("A-4")
    assert (a.start, a.done) == (at("10-02"), at("10-14"))     # the last entry into Termine, currently there
    assert a.reopened == 1 and a.excluded_from_flow_time and not a.skipped_start
    assert a.flow_time_days == 12                              # 10-14 - 10-02
    assert (a.active_days, a.wait_days) == (11, 0)             # analyse 3 + dev 4 + test 4 ; the Done interval is not Active


def test_an_issue_in_progress_has_a_start_but_no_done(row):
    a = row("A-5")
    assert (a.start, a.first_start, a.done, a.flow_time_days, a.active_days) == (at("10-12"), at("10-12"), None, None, None)


def test_a_cancelled_issue_has_a_removed_date_and_a_pending_restart(row):
    a = row("A-6")
    assert a.removed == at("10-09") and a.done is None
    assert a.first_start == at("10-03") and a.start is None    # it went back to backlog and never restarted
    assert a.backlog_returns == 1                              # Deprioritise on 10-07


def test_an_unlisted_status_is_unclassified_and_counted_apart(row):
    b = row("B-2")
    assert (b.flow_time_days, b.active_days, b.wait_days, b.unclassified_days) == (9, 6, 0, 3)
    assert {i.category for i in b.intervals} == {"Backlog", "Active", "Unclassified", "Done"}


def test_the_feature_workflow_has_its_own_statuses(row):
    p = row("PRG-1")
    assert (p.flow_time_days, p.active_days, p.wait_days) == (19, 16, 3)   # En test / recette is Active in Feature


def test_changes_given_out_of_order_give_the_same_row(flow_cases, rules):
    shuffled = flow_cases[1].__class__(**{**flow_cases[1].__dict__, "changes": tuple(reversed(flow_cases[1].changes))})
    assert item_flow(shuffled, FLOW_AS_OF, rules) == item_flow(flow_cases[1], FLOW_AS_OF, rules)


# --- Flow Load: end-of-day snapshots in UTC, half-open intervals -----------------------------------

def load(items, mode, team, period):
    return daily_load(items, mode, team, period.start, period.end, FLOW_AS_OF)


def alpha_p1(items, mode):
    return list(load(items, mode, "Alpha", P1).values())


def test_flow_load_mode_a_day_by_day(items):
    # D = 10-05 ... 10-18, snapshot at midnight of D + 1. A-1 counts on D 10-05..10-11, A-2 on 10-08..10-16 (not while
    # in backlog 10-06..10-09), A-3 never (no start), A-4 on 10-05..10-07 and 10-09..10-12 (not while Done 10-09),
    # A-5 on 10-11..10-18, A-6 on 10-05 only.
    assert alpha_p1(items, "A") == [3, 2, 2, 2, 3, 3, 4, 3, 2, 2, 2, 2, 1, 1]


def test_flow_load_mode_b_counts_from_the_last_start_only(items):
    # Mode B: A-4 counts through its Done interval (10-09), A-6 never (no start): 10-05 gives 2 (A: 3), 10-08 gives 3 (A: 2).
    assert alpha_p1(items, "B") == [2, 2, 2, 3, 3, 3, 4, 3, 2, 2, 2, 2, 1, 1]


def test_mode_a_counts_the_time_before_a_return_to_backlog_and_mode_b_does_not(items):
    a2 = next(i for i in items if i.key == "A-2")
    assert in_process_at(a2, at("10-05"), "A") is True       # 10-05 is before A-2 went back to backlog
    assert in_process_at(a2, at("10-05"), "B") is False      # its last start is 10-09
    assert in_process_at(a2, at("10-07"), "A") is False      # in Deprioritise
    assert in_process_at(a2, at("10-10"), "A") is in_process_at(a2, at("10-10"), "B") is True


def test_the_boundary_belongs_to_the_next_interval(items):
    a1 = next(i for i in items if i.key == "A-1")            # leaves A deployer and enters Termine at 10-13 00:00
    assert in_process_at(a1, dt.datetime(2026, 10, 12, 23, 59, 59), "A") is True
    assert in_process_at(a1, at("10-13"), "A") is False


def test_the_flow_load_mode_must_be_a_or_b(items):
    with pytest.raises(ValueError, match="FlowLoadMode must be A or B"):
        in_process_at(items[0], at("10-06"), "C")


def test_snapshot_days_stop_before_the_as_of_instant():
    assert len(snapshot_days(at("10-19"), at("11-02"), FLOW_AS_OF)) == 14
    assert len(snapshot_days(at("10-19"), at("11-09"), FLOW_AS_OF)) == 14                    # the future has no snapshot
    assert len(snapshot_days(at("10-19"), at("11-09"), dt.datetime(2026, 11, 1, 12))) == 13  # 11-01 has not ended


# --- the six metrics --------------------------------------------------------------------------------

def metrics(items, team, period, mode="A"):
    return team_period_metrics(items, team, period, mode, FLOW_AS_OF)


def test_alpha_first_iteration(items):
    m = metrics(items, "Alpha", P1)
    assert m.velocity == 4                                     # A-1 (10-13), A-2 (10-18), A-3 (10-08), A-4 (10-14)
    assert (m.flow_time_count, m.flow_time_excluded) == (2, 2)  # A-3 skipped start, A-4 reopened
    assert m.flow_time_median == 10                            # [9, 11]
    assert m.flow_time_p85 == pytest.approx(10.7)              # 9 + 0.85 * (11 - 9)
    assert m.flow_efficiency == pytest.approx(0.75)            # active 8 + 7 = 15 ; wait 3 + 2 = 5 ; 15 / 20
    assert m.flow_load_average == pytest.approx(32 / 14) and m.flow_load_end == 1
    assert m.distribution_done == {"Feature": 0.25, "Enabler": 0.25, "Defect": 0.25, "Autre": 0.25, "Unmapped": 0}
    assert m.distribution_wip == {"Feature": 1.0, "Enabler": 0, "Defect": 0, "Autre": 0, "Unmapped": 0}   # A-5 only
    assert m.flow_predictability == NOT_COMPUTABLE


def test_a_period_with_nothing_done_is_not_computable_not_zero(items):
    m = metrics(items, "Alpha", P2)
    assert m.velocity == 0                                     # zero is a real count
    assert m.flow_time_median == m.flow_time_p85 == m.flow_efficiency == NOT_COMPUTABLE
    assert m.distribution_done == NOT_COMPUTABLE
    assert (m.flow_load_average, m.flow_load_end) == (1, 1)    # A-5 is in process every day
    assert m.flow_time_count == m.flow_time_excluded == 0


def test_beta_second_iteration_with_an_unclassified_status(items):
    m = metrics(items, "Beta", P2)
    assert m.velocity == 2 and (m.flow_time_count, m.flow_time_excluded) == (2, 0)
    assert m.flow_time_median == 13                            # [9, 17]
    assert m.flow_time_p85 == pytest.approx(15.8)              # 9 + 0.85 * 8
    assert m.flow_efficiency == NOT_COMPUTABLE                 # B-2 has 3 days Unclassified: never a wrong figure
    assert m.distribution_done == {"Feature": 0.5, "Enabler": 0, "Defect": 0, "Autre": 0, "Unmapped": 0.5}   # B-2 has no type
    # D = 10-19 ... 11-01 (snapshot T = D + 1 = 10-20 ... 11-02): B-1 counts while T < 10-28 (8 days: T 10-20..10-27),
    # B-2 while T < 10-22 (2 days: T 10-20 and 10-21): 10 issue-days over 14 days.
    assert m.flow_load_average == pytest.approx(10 / 14) and m.flow_load_end == 0
    assert m.distribution_wip == NOT_COMPUTABLE                # nothing in process at the end


def test_beta_first_iteration_load_before_anything_is_done(items):
    m = metrics(items, "Beta", P1)
    assert m.velocity == 0
    assert m.flow_load_average == pytest.approx(16 / 14) and m.flow_load_end == 2


def test_program_features_are_reported_apart_under_art(items):
    m1, m2 = metrics(items, "ART", P1), metrics(items, "ART", P2)
    assert (m1.velocity, m1.flow_load_average, m1.flow_load_end) == (0, 1, 1)
    assert m2.velocity == 1 and m2.flow_time_median == m2.flow_time_p85 == 19
    assert m2.flow_efficiency == pytest.approx(16 / 19)
    assert m2.distribution_done["Feature"] == 1
    assert m2.flow_load_average == pytest.approx(1 / 14) and m2.flow_load_end == 0


def test_flow_metrics_lists_every_period_and_team(items):
    result = flow_metrics(items, [P1, P2], "A", FLOW_AS_OF)
    assert [(m.period, m.team) for m in result] == [
        ("P1", "ART"), ("P1", "Alpha"), ("P1", "Beta"), ("P2", "ART"), ("P2", "Alpha"), ("P2", "Beta")]


def test_a_period_after_the_as_of_date_has_no_load_figure(items):
    future = Period("F", at("11-02"), at("11-16"))
    m = metrics(items, "Alpha", future)
    assert m.flow_load_average == m.flow_load_end == m.distribution_wip == NOT_COMPUTABLE


def test_periods_are_half_open():
    assert P1.contains(at("10-05")) and not P1.contains(at("10-19")) and P2.contains(at("10-19"))
    assert not P1.contains(None)


# --- percentile = Excel PERCENTILE.INC --------------------------------------------------------------

@pytest.mark.parametrize("values,p,expected", [
    ([5], 0.85, 5), ([9, 11], 0.5, 10), ([9, 11], 0.85, 10.7), ([1, 2, 3, 4], 0.85, 3.55),
    ([4, 1, 3, 2], 0.5, 2.5), ([7, 7, 7], 0.85, 7), ([0, 10], 1.0, 10), ([0, 10], 0.0, 0)])
def test_percentile_inc(values, p, expected):
    assert percentile_inc(values, p) == pytest.approx(expected)
