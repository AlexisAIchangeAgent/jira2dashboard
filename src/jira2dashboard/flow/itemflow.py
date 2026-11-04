"""ItemFlow: one row per issue (docs/safe-flow-metrics-and-data-checks.md section 2.1)."""
import datetime as dt
from dataclasses import dataclass

from jira2dashboard.flow.intervals import build_intervals
from jira2dashboard.flow.model import (ACTIVE, BACKLOG, DONE, REMOVED, UNCLASSIFIED, WAIT, FlowIssue, Interval,
                                       Rules, days_between)


@dataclass(frozen=True)
class ItemFlow:
    key: str
    team: str
    type: str
    work_type: str | None
    created: dt.datetime
    start: dt.datetime | None            # StartUtc: where the flow time starts
    first_start: dt.datetime | None      # first entry into a start status, ever (Flow Load mode A)
    backlog_returns: int
    done: dt.datetime | None             # DoneUtc
    removed: dt.datetime | None          # RemovedUtc
    flow_time_days: float | None
    lead_time_days: float | None
    active_days: float | None
    wait_days: float | None
    unclassified_days: float | None
    reopened: int
    skipped_start: bool
    intervals: tuple[Interval, ...]

    @property
    def excluded_from_flow_time(self) -> bool:
        """Done issues that Flow Time and Flow Efficiency leave out: no valid start, or reopened after being done."""
        return self.skipped_start or self.reopened > 0


def item_flow(issue: FlowIssue, as_of: dt.datetime, rules: Rules) -> ItemFlow:
    """Compute the row of one issue.

    * StartUtc: the first entry into a status with IsStart = 1 that is strictly after the last entry into a
      Backlog status (the creation counts: the first interval of a new issue is a Backlog entry). A return
      through Backlog therefore restarts the flow time at the next entry into a start status.
    * first_start: the first entry into an IsStart status, ever (it ignores returns through Backlog).
    * BacklogReturns: entries into a Backlog status after first_start.
    * DoneUtc: the last entry into an IsEnd status, only if the issue is in an IsEnd status now. RemovedUtc:
      the same for the Removed category.
    * FlowTime = Done - Start; LeadTime = Done - Created; ActiveDays, WaitDays and UnclassifiedDays add up the
      intervals of that category entered from Start (included) to Done (excluded); all are empty without a
      Start or without a Done.
    * Reopened: transitions from the Done category to another category. SkippedStart: done with no Start.
    """
    intervals = build_intervals(issue, as_of, rules)
    workflow = issue.workflow
    starts = [i.entered for i in intervals if rules.rule(workflow, i.status).is_start]
    backlog_entries = [i.entered for i in intervals if i.category == BACKLOG]
    last_backlog = max(backlog_entries) if backlog_entries else None
    restarts = [t for t in starts if last_backlog is None or t > last_backlog]
    start = min(restarts) if restarts else None
    first_start = min(starts) if starts else None
    back = sum(1 for t in backlog_entries if first_start is not None and t > first_start)

    current = intervals[-1]
    ends = [i.entered for i in intervals if rules.rule(workflow, i.status).is_end]
    done = max(ends) if ends and rules.rule(workflow, current.status).is_end else None
    gone = [i.entered for i in intervals if i.category == REMOVED]
    removed = max(gone) if gone and current.category == REMOVED else None

    flow_time = days_between(start, done) if start is not None and done is not None else None
    lead_time = days_between(issue.created, done) if done is not None else None

    def days(category: str) -> float | None:
        if start is None or done is None:
            return None
        return sum(i.duration_days for i in intervals if i.category == category and start <= i.entered < done)

    reopened = sum(1 for before, after in zip(intervals, intervals[1:]) if before.category == DONE and after.category != DONE)
    return ItemFlow(
        key=issue.key, team=issue.team, type=issue.type, work_type=issue.work_type, created=issue.created,
        start=start, first_start=first_start, backlog_returns=back, done=done, removed=removed,
        flow_time_days=flow_time, lead_time_days=lead_time,
        active_days=days(ACTIVE), wait_days=days(WAIT), unclassified_days=days(UNCLASSIFIED),
        reopened=reopened, skipped_start=done is not None and start is None, intervals=intervals)
