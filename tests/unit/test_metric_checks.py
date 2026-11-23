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


# --- FlowEfficiencyCheck: a small workflow typed here (Nouveau backlog, Analyse active and start, Attente wait, Fini done and end) ---------

def rule(category, start=False, end=False):
    return StatusRule(category, start, end, False)


WAIT_RULES = Rules({("Feature", "Nouveau"): rule("Backlog"), ("Feature", "Analyse"): rule("Active", start=True),
                    ("Feature", "Attente"): rule("Wait"), ("Feature", "Fini"): rule("Done", end=True)})
NO_WAIT_RULES = Rules({**WAIT_RULES._rows, ("Feature", "Attente"): rule("Active")}, waits=False)
EFF = [ticket("E-1", [("2025-05-02T00:00", "Analyse"), ("2025-05-03T00:00", "Attente"), ("2025-05-04T00:00", "Analyse"), ("2025-05-06T00:00", "Fini")]),
       ticket("E-2", [("2025-05-02T00:00", "Analyse"), ("2025-05-03T00:00", "Fini")])]          # E-1: 3 active days, 1 wait; E-2: 1 active day


def efficiency(issues=EFF, rules=WAIT_RULES, iteration=I1, metric=None):
    items = [item_flow(i, AS_OF, rules) for i in issues]
    metric = metric or team_iteration_metrics(items, "ART", iteration, "A", AS_OF, rules.waits)
    return MC.flow_efficiency_check(items, iteration, "ART", rules, metric)


def test_the_active_and_wait_days_of_the_counted_items_give_the_efficiency_of_metrics():
    assert efficiency() == MC.FlowEfficiencyCheck("I1", "ART", counted=2, with_unclassified=0, active_days=4.0, wait_days=1.0, wait_statuses=1,
                                                  efficiency=0.8, gap=0.0, reason="")


def test_without_a_wait_status_the_reason_says_to_declare_one():
    got = efficiency(rules=NO_WAIT_RULES)
    assert (got.wait_statuses, got.active_days, got.wait_days, got.efficiency, got.gap, got.reason) == (0, 5.0, 0.0, 1.0, 0.0, MC.NO_WAIT)


def test_time_in_a_status_no_rule_knows_is_counted_and_said():
    odd = ticket("E-3", [("2025-05-02T00:00", "Analyse"), ("2025-05-03T00:00", "Mystere"), ("2025-05-05T00:00", "Fini")])
    got = efficiency(EFF + [odd])
    assert (got.counted, got.with_unclassified, got.reason) == (3, 1, "1 items spent time in statuses no rule knows (Unclassified)")


def test_nothing_counted_and_a_wrong_efficiency_of_metrics():
    assert efficiency(iteration=EMPTY).reason == MC.NOTHING_COUNTED
    items = [item_flow(i, AS_OF, WAIT_RULES) for i in EFF]
    right = team_iteration_metrics(items, "ART", I1, "A", AS_OF, True)
    assert efficiency(metric=right.__class__(**{**right.__dict__, "flow_efficiency": 0.5})).gap == 0.3


# --- FlowLoadCheck: the tickets EFF (E-1 in process from 05-02 to 05-06: Analyse, Attente, Analyse; E-2 from 05-02 to 05-03) ---------------
# Mode A, a snapshot at midnight of D + 1, the entry included: D = 05-01 to 05-04 hold E-1 (Fini at 05-06 00:00), D = 05-01 holds E-2 (Fini at
# 05-03 00:00): 5 item-days over the 10 days of I1.

def load(rules=WAIT_RULES, iteration=I1, as_of=AS_OF, metric=None):
    items = [item_flow(i, as_of, rules) for i in EFF]
    metric = metric or team_iteration_metrics(items, "ART", iteration, "A", as_of, rules.waits, rules.has_start("Feature"))
    return MC.flow_load_check(items, iteration, "ART", rules, "A", as_of, metric)


def test_the_days_the_snapshots_and_the_load_recounted_give_the_figures_of_metrics():
    assert load() == MC.FlowLoadCheck("I1", "ART", days=10, days_measured=10, start_statuses=1, average=0.5, at_end=0, gap=0.0, reason="")


def test_no_start_status_and_no_day_before_asof_are_the_two_reasons():
    no_start = Rules({**WAIT_RULES._rows, ("Feature", "Analyse"): rule("Active")})
    assert load(no_start) == MC.FlowLoadCheck("I1", "ART", 10, 10, 0, None, None, 0.0, MC.NO_START_STATUS)
    assert load(as_of=dt.datetime(2025, 5, 1, 12)).reason == MC.AFTER_AS_OF                  # 05-01 has not ended: no snapshot
    assert load(as_of=dt.datetime(2025, 5, 4)).days_measured == 3                              # 05-01, 05-02, 05-03


def test_a_wrong_load_at_the_end_of_metrics_shows_in_the_gap():
    items = [item_flow(i, AS_OF, WAIT_RULES) for i in EFF]
    right = team_iteration_metrics(items, "ART", I1, "A", AS_OF)
    assert load(metric=right.__class__(**{**right.__dict__, "flow_load_end": 2})).gap == -2.0



# --- the buckets of Flow Distribution, and FlowDistCheck --------------------------------------------------------------------------------

def test_a_work_type_is_known_whatever_its_case_and_keeps_its_canonical_name():
    from jira2dashboard.flow.model import work_type_bucket
    assert [work_type_bucket(w) for w in ("Feature", "feature", "DEFECT", "Autre", "Featur", "", None, 5, True)] == [
        "Feature", "Feature", "Defect", "Autre", "Unmapped", "Unmapped", "Unmapped", "Unmapped", "Unmapped"]


def typed(key, work_type):
    t = ticket(key, [("2025-05-02T00:00", "En analyse"), ("2025-05-06T00:00", "Terminé")])
    return t.__class__(**{**t.__dict__, "work_type": work_type})


def distribution(issues, field="customfield_1"):
    items = [item_flow(i, AS_OF, RULES) for i in issues]
    metric = team_iteration_metrics(items, "ART", I1, "A", AS_OF)
    return MC.flow_distribution_check(items, I1, "ART", metric, field)


def test_the_done_items_are_counted_again_per_work_type_and_match_the_shares_of_metrics():
    got = distribution([typed("D-1", "Feature"), typed("D-2", "defect"), typed("D-3", "Feature"), typed("D-4", "Story")])
    assert got == MC.FlowDistCheck("I1", "ART", done=4, counts=(2, 0, 1, 0, 1), gap=0.0, reason="", note="")


def test_every_done_item_unmapped_says_why_and_nothing_done_is_not_computable():
    assert distribution([typed("D-1", None)], field="").note == "every done item is Unmapped: WorkTypeField is empty in Config"
    assert distribution([typed("D-1", None)], field="\u00a0\t\n").note.endswith(MC.NO_WORK_TYPE_FIELD)     # blanks the formula leaves out
    assert distribution([typed("D-1", "Story")]).note == "every done item is Unmapped: no known work type (Feature, Enabler, Defect, Autre)"
    items = []
    assert MC.flow_distribution_check(items, I1, "ART", team_iteration_metrics(items, "ART", I1, "A", AS_OF), "x").reason == MC.NO_ITEM
