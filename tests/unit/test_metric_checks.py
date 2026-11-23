"""The accounting of the flow metrics (flow/metric_checks.py), iteration by iteration and team by team (AL, 2026-11-22). Every figure is typed by hand
from the small tickets below (iteration I1 = 05-01 .. 05-11, End excluded; the Feature workflow of params/workflows.md: Nouveau, En analyse = start, ...,
Terminé = end).
"""
import datetime as dt

from jira2dashboard.flow import metric_checks as MC
from jira2dashboard.flow.itemflow import item_flow
from jira2dashboard.flow.metrics import team_iteration_metrics
from jira2dashboard.flow.model import Change, FlowIssue, Iteration, Rules, StatusRule

I1 = Iteration("I1", dt.datetime(2025, 5, 1), dt.datetime(2025, 5, 11))
EMPTY = Iteration("I9", dt.datetime(2025, 9, 1), dt.datetime(2025, 9, 11))
RULES = Rules.from_params()
NO_START = Rules({k: StatusRule(r.category, False, r.is_end, r.is_free) for k, r in RULES._rows.items()})
AS_OF = dt.datetime(2025, 12, 1)


def ticket(key, moves, team="ART"):
    changes, previous = [], "Nouveau"
    for n, (at, status) in enumerate(moves, start=1):
        changes.append(Change(f"{key}-{n}", dt.datetime.fromisoformat(at), previous, status))
        previous = status
    return FlowIssue(key=key, team=team, type="Feature", work_type="Feature", workflow="Feature", created=dt.datetime(2025, 4, 1),
                     status=previous, changes=tuple(changes))


ISSUES = [
    ticket("A-1", [("2025-05-02T00:00", "En analyse"), ("2025-05-06T00:00", "Terminé")]),                         # counted
    ticket("A-2", [("2025-05-03T00:00", "En analyse"), ("2025-05-07T00:00", "Terminé")]),                         # counted
    ticket("R-1", [("2025-05-02T00:00", "En analyse"), ("2025-05-03T00:00", "Terminé"), ("2025-05-04T00:00", "En développement"),
                   ("2025-05-06T00:00", "Terminé")]),                                                               # reopened
    ticket("S-1", [("2025-05-06T00:00", "Terminé")]),                                                               # Nouveau -> Terminé: no start
    ticket("B-1", [("2025-05-02T00:00", "En analyse"), ("2025-05-06T00:00", "Terminé")], team="Beta"),            # another team
    ticket("O-1", [("2025-05-02T00:00", "En analyse")]),                                                            # open
]


def check(rules=RULES, iteration=I1, team="ART", metric=None):
    items = [item_flow(i, AS_OF, rules) for i in ISSUES]
    metric = metric or team_iteration_metrics(items, team, iteration, "A", AS_OF)
    return MC.flow_time_check(ISSUES, items, iteration, team, rules, metric)


def test_the_done_items_are_the_counted_ones_and_the_left_out_ones_each_under_one_reason():
    assert check() == MC.FlowTimeCheck("I1", "ART", done=4, excl_reopened=1, excl_no_start_rule=0, excl_not_started=1, counted=2, gap=0, reason="")


def test_without_a_start_status_in_the_rules_every_done_item_is_left_out_and_the_reason_says_what_to_set():
    got = check(NO_START)
    assert (got.done, got.excl_reopened, got.excl_no_start_rule, got.excl_not_started, got.counted, got.gap) == (4, 1, 3, 0, 0, 0)
    assert got.reason == "all 4 done items left out: 3 no start status in the rules (set IsStart in Workflows), 1 reopened"


def test_an_iteration_with_nothing_done_says_so():
    assert check(iteration=EMPTY) == MC.FlowTimeCheck("I9", "ART", 0, 0, 0, 0, 0, 0, "no item done in the iteration")


def test_another_team_is_counted_apart():
    assert check(team="Beta") == MC.FlowTimeCheck("I1", "Beta", 1, 0, 0, 0, 1, 0, "")


def test_a_wrong_count_of_the_metric_shows_in_the_gap():
    """The Gap compares two paths: the done items counted from the intervals minus the left out ones, and the count of Metrics."""
    items = [item_flow(i, AS_OF, RULES) for i in ISSUES]
    right = team_iteration_metrics(items, "ART", I1, "A", AS_OF)
    wrong = right.__class__(**{**right.__dict__, "flow_time_count": 3})
    assert check(metric=wrong).gap == -1


def test_the_reason_lists_the_causes_in_their_order_and_only_the_ones_that_apply():
    assert MC.flow_time_reason(5, 2, 0, 3, 0) == "all 5 done items left out: 3 no start after the backlog, 2 reopened"
    assert MC.flow_time_reason(5, 2, 0, 3, 1) == "" and MC.flow_time_reason(0, 0, 0, 0, 0) == MC.NO_ITEM
