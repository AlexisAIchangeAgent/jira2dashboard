"""The accounting of each flow metric, iteration by iteration and team by team (AL, 2026-11-22: "flow metrics are still not computable, I need the same
kind of accountability checks made for predictability for each flow metric").

One table per metric, the keys of Metrics (iteration x team), like VelocityCheck and IterationFeatures: a funnel of counts from the population to the
figure, a Gap that compares two independent paths and is 0, and a Reason that says in clear why a point is not computable ("" when it is). Excel: the
tables of formulas.FLOWTIMECHECK_COLUMNS...; the texts of the reasons are these constants on both sides.

FlowTimeCheck: Done is counted again from the status intervals (an open end status entered in the iteration, as VelocityCheck counts the changes); the
done items left out of Flow Time each under ONE reason, the first that applies: reopened, then no start status in the rules of their workflow (a setting
to fix: IsStart in the sheet Workflows), then no start after their last backlog entry; Counted is Metrics.FlowTimeCount. Gap = Done - the three - Counted.
"""
from dataclasses import dataclass

from jira2dashboard.flow.itemflow import ItemFlow
from jira2dashboard.flow.metrics import TeamIterationMetrics
from jira2dashboard.flow.model import ACTIVE, WAIT, FlowIssue, Iteration, Rules

NO_ITEM = "no item done in the iteration"
ALL, LEFT_OUT = "all ", " done items left out: "
NO_START_RULE = " no start status in the rules (set IsStart in Workflows)"
NOT_STARTED = " no start after the backlog"
REOPENED = " reopened"


@dataclass(frozen=True)
class FlowTimeCheck:
    iteration: str
    team: str
    done: int                    # done in the iteration, counted from the intervals
    excl_reopened: int           # done, left out of Flow Time because reopened
    excl_no_start_rule: int      # not reopened, left out because the workflow has no start status
    excl_not_started: int        # not reopened, left out because it never started after its last backlog entry
    counted: int                 # Metrics.FlowTimeCount
    gap: int
    reason: str


def flow_time_reason(done: int, reopened: int, no_start_rule: int, not_started: int, counted: int) -> str:
    """Why the Flow Time of the point is not computable; "" when something is counted. Whole counts only (a decimal reads 1,5 in a French Excel)."""
    if counted > 0:
        return ""
    if done == 0:
        return NO_ITEM
    parts = [f"{n}{text}" for n, text in ((no_start_rule, NO_START_RULE), (not_started, NOT_STARTED), (reopened, REOPENED)) if n > 0]
    return f"{ALL}{done}{LEFT_OUT}{', '.join(parts)}"


def flow_time_check(issues: list[FlowIssue], items: list[ItemFlow], iteration: Iteration, team: str, rules: Rules,
                    metric: TeamIterationMetrics) -> FlowTimeCheck:
    """One row: `items[n]` is the ItemFlow of `issues[n]`, `metric` the row of Metrics of the same iteration and team."""
    done = sum(1 for issue, item in zip(issues, items) if item.team == team and any(
        visit.left is None and rules.rule(issue.workflow, visit.status).is_end and iteration.contains(visit.entered) for visit in item.intervals))
    finished = [item for item in items if item.team == team and iteration.contains(item.done)]
    reopened = sum(1 for item in finished if item.reopened > 0)
    no_start_rule = sum(1 for item in finished if item.reopened == 0 and item.skipped_start and not item.start_rule)
    not_started = sum(1 for item in finished if item.reopened == 0 and item.skipped_start and item.start_rule)
    counted = metric.flow_time_count
    return FlowTimeCheck(iteration.name, team, done, reopened, no_start_rule, not_started, counted,
                         done - reopened - no_start_rule - not_started - counted, flow_time_reason(done, reopened, no_start_rule, not_started, counted))


# --- FlowEfficiencyCheck: Counted (as Flow Time), the counted items that spent time in a status no rule knows, the Active and Wait days summed
# again from the status intervals (Excel: StatusIntervals.InCycle, an independent path from ItemFlow.ActiveDays / WaitDays), the Wait statuses of
# the rules, the efficiency recounted, Gap = recount - Metrics.FlowEfficiency where it is computable (0 elsewhere), and the Reason in the order of the
# formula of Metrics: no Wait status, nothing counted, unclassified time, no active or wait time.
NO_WAIT = "no status is declared Wait (set Category = Wait in Workflows)"
NOTHING_COUNTED = "no item counted in Flow Time (see FlowTimeCheck)"
UNCLASSIFIED_TIME = " items spent time in statuses no rule knows (Unclassified)"
NO_TIME = "no active or wait time"
GAP_DIGITS = 9                     # the Gap is rounded: two sums of days in another order differ in the last bits


@dataclass(frozen=True)
class FlowEfficiencyCheck:
    iteration: str
    team: str
    counted: int                 # the items behind Flow Time and Flow Efficiency (Metrics.FlowTimeCount)
    with_unclassified: int       # of which with time in an Unclassified status
    active_days: float           # summed from the intervals of the counted items, from their start (included) to their end (excluded)
    wait_days: float
    wait_statuses: int           # rules whose category is Wait (both workflow kinds)
    efficiency: float | None     # active / (active + wait); None without any time
    gap: float
    reason: str


def flow_efficiency_reason(wait_statuses: int, counted: int, with_unclassified: int, active: float, wait: float) -> str:
    if wait_statuses == 0:
        return NO_WAIT
    if counted == 0:
        return NOTHING_COUNTED
    if with_unclassified > 0:
        return f"{with_unclassified}{UNCLASSIFIED_TIME}"
    return NO_TIME if active + wait == 0 else ""


def flow_efficiency_check(items: list[ItemFlow], iteration: Iteration, team: str, rules: Rules, metric: TeamIterationMetrics) -> FlowEfficiencyCheck:
    """One row, `metric` the row of Metrics of the same iteration and team."""
    counted = [i for i in items if i.team == team and iteration.contains(i.done) and not i.excluded_from_flow_time]

    def days(category: str) -> float:
        return sum(v.duration_days for i in counted for v in i.intervals if v.category == category and i.start <= v.entered < i.done)
    active, wait = days(ACTIVE), days(WAIT)
    with_unclassified = sum(1 for i in counted if (i.unclassified_days or 0) > 0)
    waits = sum(1 for r in rules.listed() if r.category == WAIT)
    efficiency = active / (active + wait) if active + wait > 0 else None
    reason = flow_efficiency_reason(waits, len(counted), with_unclassified, active, wait)
    if reason:
        gap = 0.0
    elif isinstance(metric.flow_efficiency, float) and efficiency is not None:
        gap = round(efficiency - metric.flow_efficiency, GAP_DIGITS)
    else:
        gap = 1.0                                                   # computable here, not in Metrics: a difference to look at
    return FlowEfficiencyCheck(iteration.name, team, len(counted), with_unclassified, active, wait, waits, efficiency, gap, reason)
