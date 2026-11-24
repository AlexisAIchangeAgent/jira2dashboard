"""The reconciliation: the controls an accountant runs on the ledger of the figures (AL, 2026-11-18, "c'est de la comptabilite").

Each control is (name, assertion, expected, found) and passes when found == expected (a count to REVIEW, such as the cut-off one, expects 0):
the audit assertions are completeness (nothing lost between the extraction and the figures), accuracy (a total recounted from the ledger), occurrence
(no ticket twice), cut-off (the tickets near an iteration boundary), classification (the statuses and the Feature types are the right ones) and
existence (the done tickets exist as an end status in the changelog). The ledger is the sheet MetricItems (flow/trace.metric_items).
Excel: the table Reconciliation (formulas.RECONCILIATION_*), the census is the block of spilled formulas next to it.
"""
from collections import Counter
from dataclasses import dataclass

from jira2dashboard.flow.itemflow import ItemFlow
from jira2dashboard.flow.iteration_features import IterationFeatures, carry_over_gap
from jira2dashboard.flow.metric_checks import flow_distribution_check, flow_efficiency_check, flow_load_check, flow_time_check
from jira2dashboard.flow.metrics import TeamIterationMetrics
from jira2dashboard.flow.model import ART, OTHER_TYPES, UNCLASSIFIED, WAIT, FlowIssue, Iteration, Rules
from jira2dashboard.flow.trace import ACCOUNT_COUNTED, ACCOUNT_OTHER, ACCOUNTS, CUTOFF_HOURS, MetricItems, account, metric_items, trace

POPULATION = "Tickets read from Jira = tickets of ItemFlow"
ACCOUNTS_ADD_UP = "Every ticket sits in one account"
VELOCITY_TOTAL = "Velocity of all iterations = tickets of the Done list"
VELOCITY_ROWS = "Velocity of each iteration and team recounted from the list (largest difference)"
FLOW_TIME_COUNT = "Flow Time count = tickets counted in the Done list"
FLOW_TIME_DAYS = "Flow Time days recounted from the list"
FLOW_TIME_GAP = "Flow Time counted = done minus the left out, every iteration and team (FlowTimeCheck)"
FLOW_EFFICIENCY_GAP = "Flow Efficiency recounted from the status intervals, every iteration and team (FlowEfficiencyCheck)"
FLOW_LOAD_GAP = "Flow Load recounted at the end of each iteration, every team (FlowLoadCheck)"
FLOW_DISTRIBUTION_GAP = "Flow Distribution recounted per work type, every iteration and team (FlowDistCheck)"
CARRY_OVER = "What an iteration carries over = what the next one of the PI engages (IterationFeatures)"
UNIQUE = "No ticket twice in the Done list"
CUTOFF = f"Tickets done within {CUTOFF_HOURS} hours of an iteration boundary (to review)"
CLASSIFIED = "Statuses of the changelog of the counted tickets all classified in Workflows"
FEATURE_TYPES = "FeatureTypes matches at least one ticket of the Program"
# AL, 2026-11-22: the two settings without which Flow Time, Flow Efficiency and Flow Load cannot be computed (no control said it: every row stayed OK)
START_STATUS = "Every workflow kind with tickets has a start status (IsStart)"
WAIT_STATUS = "A status is declared Wait (Flow Efficiency)"
HISTORIES = "Status histories consistent (DC06, DC07)"
EXISTENCE = "Done tickets of ItemFlow = open end-status intervals of the changelog"
# (name, assertion), in the order of the sheet
CONTROLS = ((POPULATION, "Completeness"), (ACCOUNTS_ADD_UP, "Completeness"), (VELOCITY_TOTAL, "Accuracy"), (VELOCITY_ROWS, "Accuracy"),
            (FLOW_TIME_COUNT, "Accuracy"), (FLOW_TIME_DAYS, "Accuracy"), (FLOW_TIME_GAP, "Accuracy"), (FLOW_EFFICIENCY_GAP, "Accuracy"),
            (FLOW_LOAD_GAP, "Accuracy"), (FLOW_DISTRIBUTION_GAP, "Accuracy"), (CARRY_OVER, "Completeness"), (UNIQUE, "Occurrence"), (CUTOFF, "Cut-off"), (CLASSIFIED, "Classification"),
            (FEATURE_TYPES, "Classification"), (START_STATUS, "Classification"), (WAIT_STATUS, "Classification"), (HISTORIES, "Existence"),
            (EXISTENCE, "Existence"))
TOLERANCE = 1e-6
INFO_ASSERTIONS = ("Cut-off",)        # a control of these assertions that does not match is a point of attention (INFO), not an error to review


@dataclass(frozen=True)
class Control:
    name: str
    assertion: str
    expected: float
    found: float

    @property
    def passed(self) -> bool:
        return abs(self.found - self.expected) < TOLERANCE

    @property
    def verdict(self) -> str:
        """OK, INFO (a point of attention: the Dashboard does not count it) or REVIEW (look at it: the Dashboard counts it)."""
        return "OK" if self.passed else "INFO" if self.assertion in INFO_ASSERTIONS else "REVIEW"


def reconcile(issues: list[FlowIssue], items: list[ItemFlow], metrics: list[TeamIterationMetrics], iterations: list[Iteration],
              pis: dict, rules: Rules, flags: dict[str, frozenset[str]], lists: MetricItems | None = None, mode: str = "A",
              as_of=None, features: list[IterationFeatures] | None = None) -> list[Control]:
    """`items[n]` is the ItemFlow of `issues[n]`; `metrics` the rows of the table Metrics; `flags` the data checks per key (checks.run_checks);
    `lists` the ledger when it is not the one that metric_items builds (a test of the controls feeds a wrong one); `mode` and `as_of` those of Flow
    Load (as_of None: every day of the iterations is measured); `features` the rows of IterationFeatures (none: the control finds 0)."""
    lists = lists or metric_items(items, iterations, pis)
    counted = [r for r in lists.done if r[7] == 1]
    accounts = Counter(account(item, trace(item, iterations, pis).done_iteration != "") for item in items)
    rows = Counter((r[0], r[3]) for r in lists.done)
    program = [i for i in issues if i.workflow == "Feature"]
    by_name = {i.name: i for i in iterations}
    values = {
        POPULATION: (len(issues), len(items)),
        ACCOUNTS_ADD_UP: (len(items), sum(accounts[a] for a in ACCOUNTS)),
        VELOCITY_TOTAL: (sum(m.velocity for m in metrics), len(lists.done)),
        VELOCITY_ROWS: (0, max((abs(m.velocity - rows[(m.iteration, m.team)]) for m in metrics), default=0)),
        FLOW_TIME_COUNT: (sum(m.flow_time_count for m in metrics), len(counted)),
        FLOW_TIME_DAYS: (sum(i.flow_time_days or 0 for i in items if account(i, trace(i, iterations, pis).done_iteration != "") == ACCOUNT_COUNTED
                             and not i.excluded_from_flow_time), sum(r[8] or 0 for r in counted if r[8] != "")),
        FLOW_TIME_GAP: (0, sum(abs(flow_time_check(issues, items, by_name[m.iteration], m.team, rules, m).gap) for m in metrics)),
        FLOW_EFFICIENCY_GAP: (0, sum(abs(flow_efficiency_check(items, by_name[m.iteration], m.team, rules, m).gap) for m in metrics)),
        FLOW_LOAD_GAP: (0, sum(abs(flow_load_check(items, by_name[m.iteration], m.team, rules, mode, as_of or max(i.end for i in iterations), m).gap)
                               for m in metrics)),
        FLOW_DISTRIBUTION_GAP: (0, sum(abs(flow_distribution_check(items, by_name[m.iteration], m.team, m, "").gap) for m in metrics)),
        CARRY_OVER: (0, carry_over_gap(features or [])),
        UNIQUE: (0, len(lists.done) - len({r[4] for r in lists.done})),
        CUTOFF: (0, len(lists.cutoff)),
        CLASSIFIED: (0, sum(1 for item in items if item.team != OTHER_TYPES for visit in item.intervals if visit.category == UNCLASSIFIED)),
        FEATURE_TYPES: (1, 1 if not program else min(1, sum(1 for i in items if i.team == ART))),
        START_STATUS: (0, sum(1 for kind in ("Feature", "Team") if any(i.workflow == kind for i in issues) and not rules.has_start(kind))),
        WAIT_STATUS: (1, 1 if not issues else min(1, sum(1 for r in rules.listed() if r.category == WAIT))),
        HISTORIES: (0, sum(1 for ids in flags.values() for check in ("DC06", "DC07") if check in ids)),
        EXISTENCE: (sum(1 for i in items if i.done is not None),
                    sum(1 for issue, item in zip(issues, items) for visit in item.intervals
                        if visit.is_open and rules.rule(issue.workflow, visit.status).is_end)),
    }
    return [Control(name, assertion, *values[name]) for name, assertion in CONTROLS]


def unclassified(items: list[ItemFlow]) -> list[tuple]:
    """The statuses that no rule classifies, one row per status in the order of the names: status, intervals of the counted tickets, intervals of the
    tickets left out (the team "ART (other types)"): the first number is what the control counts, the second tells that it is not the ART's business."""
    counted, left_out = Counter(), Counter()
    for item in items:
        for visit in item.intervals:
            if visit.category == UNCLASSIFIED:
                (left_out if item.team == OTHER_TYPES else counted)[visit.status] += 1
    return [(status, counted[status], left_out[status]) for status in sorted(set(counted) | set(left_out))]


def done_statuses(issues: list[FlowIssue], rules: Rules) -> list[tuple]:
    """The statuses that end the flow and the tickets that sit in each now, one row per status in the order of the names: what the person reads to declare as
    exceptions the Done statuses of Jira that are no delivery (cancelled, to deploy): status, tickets."""
    ended = rules.ended_statuses()
    count = Counter(issue.status for issue in issues if issue.status in ended)
    return [(status, count[status]) for status in sorted(ended)]


def census(items: list[ItemFlow], iterations: list[Iteration], pis: dict) -> list[tuple]:
    """One row per ticket type, in the order of the names: type, tickets, in the team ART, left out (the other team), done, done in an iteration."""
    out = []
    for type_ in sorted({i.type for i in items}):
        mine = [i for i in items if i.type == type_]
        traces = [trace(i, iterations, pis) for i in mine]
        out.append((type_, len(mine), sum(1 for i in mine if i.team == ART), sum(1 for t in traces if t.account == ACCOUNT_OTHER),
                    sum(1 for i in mine if i.done is not None), sum(1 for t in traces if t.done_iteration != "")))
    return out
