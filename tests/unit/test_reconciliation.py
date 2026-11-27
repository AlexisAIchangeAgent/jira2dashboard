"""The reconciliation (flow/reconciliation.py): the controls an accountant runs on the ledger of the figures (AL, 2026-11-18), and the census of the
types. Every control is (name, assertion, expected, found); it passes when found == expected. Figures typed by hand from the small tickets below
(iterations I1 = 05-01 .. 05-11 and I2 = 05-11 .. 05-21, End excluded; Feature workflow of params/workflows.md).
"""
import datetime as dt

import pytest

from jira2dashboard.flow.itemflow import item_flow
from jira2dashboard.flow.metrics import team_iteration_metrics
from jira2dashboard.flow.model import OTHER_TYPES, Change, FlowIssue, Iteration, Rules, StatusRule
from jira2dashboard.flow.reconciliation import CARRY_OVER, START_STATUS, WAIT_STATUS, Control, census, reconcile, unclassified
from jira2dashboard.flow.trace import MetricItems, metric_items

I1 = Iteration("I1", dt.datetime(2025, 5, 1), dt.datetime(2025, 5, 11))
I2 = Iteration("I2", dt.datetime(2025, 5, 11), dt.datetime(2025, 5, 21))
ITERATIONS = [I1, I2]
PIS = {"P1": (dt.datetime(2025, 5, 1), dt.datetime(2025, 6, 1))}
RULES = Rules.from_params()
AS_OF = dt.datetime(2025, 12, 1)
FEATURES = frozenset({"Feature"})


def day(text):
    return dt.datetime.fromisoformat(text)


def issue(key, moves, team="ART", type_="Feature", status_from="Nouveau"):
    changes, previous = [], status_from
    for n, (at, status) in enumerate(moves, start=1):
        changes.append(Change(f"{key}-{n}", day(at), previous, status))
        previous = status
    return FlowIssue(key=key, team=team, type=type_, work_type="Feature", workflow="Feature", created=day("2025-04-01T00:00"),
                     status=previous, changes=tuple(changes))


ISSUES = [
    issue("A-1", [("2025-05-02T00:00", "En analyse"), ("2025-05-06T00:00", "Terminé")]),                                  # done in I1
    issue("W-2", [("2025-05-02T00:00", "En analyse"), ("2025-05-03T00:00", "Terminé"), ("2025-05-04T00:00", "En développement"),
                  ("2025-05-06T00:00", "Terminé")]),                                                                      # reopened: done in I1, not in the flow time
    issue("C-1", [("2025-05-09T00:00", "En analyse"), ("2025-05-10T20:00", "Terminé")]),                                  # 4 h before the end of I1
    issue("A-3", [("2025-05-02T00:00", "En analyse"), ("2025-05-25T00:00", "Terminé")]),                                  # done after every iteration
    issue("S-1", [("2025-05-02T00:00", "En analyse"), ("2025-05-06T00:00", "Terminé")], team=OTHER_TYPES, type_="Sous-tâche"),   # not a Feature
    issue("O-1", [("2025-05-02T00:00", "En analyse")]),                                                                   # open
]


def controls(issues=ISSUES, flags=None, rules=RULES):
    items = [item_flow(i, AS_OF, rules) for i in issues]
    metrics = [team_iteration_metrics(items, "ART", it, "A", AS_OF) for it in ITERATIONS]
    return {c.name: c for c in reconcile(issues, items, metrics, ITERATIONS, PIS, rules, flags or {})}


def test_the_population_is_complete_and_every_ticket_sits_in_one_account():
    got = controls()
    assert (got["Tickets read from Jira = tickets of ItemFlow"].expected, got["Tickets read from Jira = tickets of ItemFlow"].found) == (6, 6)
    assert (got["Every ticket sits in one account"].expected, got["Every ticket sits in one account"].found) == (6, 6)


def test_the_velocity_is_recounted_from_the_list_in_total_and_for_each_iteration_and_team():
    got = controls()
    # done in an iteration, team ART: A-1, W-2, C-1 (I1); S-1 is of the other team and A-3 is outside: Velocity I1 = 3, I2 = 0
    assert (got["Velocity of all iterations = tickets of the Done list"].expected, got["Velocity of all iterations = tickets of the Done list"].found) == (3, 3)
    assert got["Velocity of each iteration and team recounted from the list (largest difference)"].found == 0
    assert (got["Flow Time count = tickets counted in the Done list"].expected, got["Flow Time count = tickets counted in the Done list"].found) == (2, 2)   # A-1, C-1
    days = got["Flow Time days recounted from the list"]
    assert days.expected == pytest.approx(4.0 + 1 + 20 / 24) and days.found == pytest.approx(days.expected)               # A-1 4 days, C-1 1 d 20 h
    assert got["No ticket twice in the Done list"].found == 0


def test_the_controls_that_ask_for_a_review_count_what_to_look_at():
    got = controls()
    assert got["Tickets done within 24 hours of an iteration boundary (to review)"].found == 1                  # C-1; the other team's S-1 is not listed
    assert got["Statuses of the changelog of the counted tickets all classified in Workflows"].found == 0
    assert got["FeatureTypes matches at least one ticket of the Program"] == Control(
        "FeatureTypes matches at least one ticket of the Program", "Classification", 1, 1)
    assert got["Done tickets of ItemFlow = open end-status intervals of the changelog"].expected == 5            # A-1, W-2, C-1, A-3, S-1
    assert got["Done tickets of ItemFlow = open end-status intervals of the changelog"].found == 5


def test_a_workflow_without_a_start_status_and_no_wait_status_are_reviews():
    """AL, 2026-11-22: the settings without which Flow Time, Flow Efficiency and Flow Load cannot be computed, which no control used to tell."""
    got = controls()
    assert (got[START_STATUS].found, got[START_STATUS].verdict, got[WAIT_STATUS].found, got[WAIT_STATUS].verdict) == (0, "OK", 1, "OK")
    bare = Rules({k: StatusRule("Active" if r.category == "Wait" else r.category, False, r.is_end, r.is_free) for k, r in RULES._rows.items()})
    got = controls(rules=bare)
    assert (got[START_STATUS].expected, got[START_STATUS].found, got[START_STATUS].verdict) == (0, 1, "REVIEW")   # Feature only: no Team ticket here
    assert (got[WAIT_STATUS].expected, got[WAIT_STATUS].found, got[WAIT_STATUS].verdict) == (1, 0, "REVIEW")
    team = issue("T-1", [("2025-05-02T00:00", "En analyse")], team="Alpha")
    team = FlowIssue(**{**team.__dict__, "workflow": "Team"})
    assert controls(ISSUES + [team], rules=bare)[START_STATUS].found == 2
    assert controls([], rules=bare)[WAIT_STATUS].verdict == "OK"                                                   # no ticket, nothing to compute


def test_a_program_with_no_feature_fails_the_feature_types_control():
    only_others = [issue("S-1", [("2025-05-02T00:00", "En analyse"), ("2025-05-06T00:00", "Terminé")], team=OTHER_TYPES, type_="Sous-tâche")]
    got = controls(only_others)["FeatureTypes matches at least one ticket of the Program"]
    assert (got.expected, got.found) == (1, 0)


def test_the_status_histories_control_counts_the_tickets_flagged_dc06_or_dc07():
    got = controls(flags={"A-1": frozenset({"DC06"}), "W-2": frozenset({"DC07", "DC01"}), "C-1": frozenset({"DC10"})})
    assert got["Status histories consistent (DC06, DC07)"].found == 2


def test_a_status_missing_from_workflows_is_counted_as_unclassified():
    odd = issue("U-1", [("2025-05-02T00:00", "Inconnu")])
    assert controls([odd])["Statuses of the changelog of the counted tickets all classified in Workflows"].found == 1


def test_the_census_counts_each_type_by_team_and_account_in_the_order_of_the_names():
    items = [item_flow(i, AS_OF, RULES) for i in ISSUES]
    # type, tickets, in ART, left out, done, done in an iteration
    assert census(items, ITERATIONS, PIS) == [("Feature", 5, 5, 0, 4, 3), ("Sous-tâche", 1, 0, 1, 1, 1)]


def test_a_ticket_twice_in_the_ledger_is_found_by_the_uniqueness_control_and_the_totals():
    items = [item_flow(i, AS_OF, RULES) for i in ISSUES]
    metrics = [team_iteration_metrics(items, "ART", it, "A", AS_OF) for it in ITERATIONS]
    ledger = metric_items(items, ITERATIONS, PIS)
    wrong = MetricItems(ledger.done + [ledger.done[0]], ledger.left_out, ledger.cutoff)                  # A-1 listed twice
    got = {c.name: c for c in reconcile(ISSUES, items, metrics, ITERATIONS, PIS, RULES, {}, lists=wrong)}
    assert got["No ticket twice in the Done list"].found == 1
    assert not got["Velocity of all iterations = tickets of the Done list"].passed                         # 4 rows for a Velocity of 3
    assert not got["Velocity of each iteration and team recounted from the list (largest difference)"].passed


NAME_CLASSIFIED = "Statuses of the changelog of the counted tickets all classified in Workflows"


def test_an_unknown_status_of_a_left_out_ticket_is_not_an_error_but_one_of_a_feature_is():
    left_out = issue("S-2", [("2025-05-02T00:00", "Inconnu")], team=OTHER_TYPES, type_="Sous-tâche")
    feature = issue("U-2", [("2025-05-02T00:00", "Inconnu")])
    assert controls([left_out])[NAME_CLASSIFIED].found == 0                    # the ART does not count that ticket
    assert controls([feature])[NAME_CLASSIFIED].found == 1
    assert controls([left_out, feature])[NAME_CLASSIFIED].found == 1


def test_the_cut_off_control_is_a_point_of_attention_not_a_review_and_the_others_are_reviews():
    got = controls()
    assert got["Tickets done within 24 hours of an iteration boundary (to review)"].verdict == "INFO"          # C-1 is 4 h from a boundary
    assert got[NAME_CLASSIFIED].verdict == "OK" and got["No ticket twice in the Done list"].verdict == "OK"
    assert controls([issue("U-2", [("2025-05-02T00:00", "Inconnu")])])[NAME_CLASSIFIED].verdict == "REVIEW"
    assert Control("x", "Accuracy", 1, 2).verdict == "REVIEW" and Control("x", "Cut-off", 0, 0).verdict == "OK"


def test_the_unclassified_statuses_are_listed_by_name_with_the_intervals_of_the_counted_and_of_the_left_out_tickets():
    population = [issue("U-1", [("2025-05-02T00:00", "Inconnu"), ("2025-05-03T00:00", "Autre")]),                 # Inconnu 1 interval, Autre 1 (open)
                  issue("U-2", [("2025-05-02T00:00", "Inconnu")]),                                                  # Inconnu 1 (open)
                  issue("S-3", [("2025-05-02T00:00", "Inconnu")], team=OTHER_TYPES, type_="Sous-tâche")]            # Inconnu 1, left out
    items = [item_flow(i, AS_OF, RULES) for i in population]
    assert unclassified(items) == [("Autre", 1, 0), ("Inconnu", 2, 1)]
    assert unclassified([item_flow(i, AS_OF, RULES) for i in ISSUES]) == []


def test_the_statuses_that_end_the_flow_are_listed_with_the_tickets_that_sit_in_them_now():
    from jira2dashboard.flow.reconciliation import done_statuses
    from jira2dashboard.params.model import ParamSheet
    header = ("Workflow", "Status", "Category", "IsEnd", "IsFree")
    issues = [issue("A", [("2025-05-02T00:00", "done"), ("2025-05-03T00:00", "cancelled")]), issue("B", [("2025-05-02T00:00", "cancelled")]),
              issue("C", [("2025-05-02T00:00", "done")]), issue("D", [("2025-05-02T00:00", "doing")])]
    catalog = {"done": "DONE", "cancelled": "DONE", "doing": "IN_PROGRESS"}                                   # Jira calls both done
    rules = Rules.from_params([ParamSheet("Workflows", header, ())], catalog, {})
    assert done_statuses(issues, rules) == [("cancelled", 2), ("done", 1)]
    declared = Rules.from_params([ParamSheet("Workflows", header, ((None, "cancelled", "Removed", None, None),))], catalog, {})
    assert done_statuses(issues, declared) == [("done", 1)]                                                     # an exception takes it off the list



def test_what_an_iteration_carries_over_is_what_the_next_one_of_the_pi_engages():
    """Typed here: I1 then I2 (contiguous), both of PI-1, carry over 3 then engage 2: a gap of 1. I3 starts a day after I2 ends (not judged), I2b of
    another PI starts where I2 ends (not judged), the other team is apart."""
    import datetime as dt
    from jira2dashboard.flow.iteration_features import IterationFeatures, carry_over_gap

    def row(name, start, end, carry, engaged, pi="PI-1", team="ART"):
        return IterationFeatures(name, dt.datetime(2026, 1, start), dt.datetime(2026, 1, end), team, pi, None, engaged, 0, 0, 0, 0, 0, carry)
    rows = [row("I1", 1, 11, carry=3, engaged=5), row("I2", 11, 21, carry=4, engaged=2), row("I3", 22, 31, carry=0, engaged=9),
            row("I2b", 21, 31, carry=0, engaged=7, pi="PI-2"), row("I1", 1, 11, carry=1, engaged=1, team="B"), row("I2", 11, 21, carry=0, engaged=1, team="B")]
    assert carry_over_gap(rows) == 1
    assert carry_over_gap(rows[:1] + rows[2:]) == 0
    issues = ISSUES
    items = [item_flow(i, AS_OF, RULES) for i in issues]
    metrics = [team_iteration_metrics(items, "ART", it, "A", AS_OF) for it in ITERATIONS]
    got = {c.name: c for c in reconcile(issues, items, metrics, ITERATIONS, PIS, RULES, {}, features=rows)}
    assert got[CARRY_OVER].found == 1 and got[CARRY_OVER].verdict == "REVIEW"
