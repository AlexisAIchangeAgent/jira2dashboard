"""StatusIntervals: the status history of an issue as intervals. Mirrors modFlow.BuildStatusIntervals."""
import datetime as dt

from jira2dashboard.flow.model import FlowIssue, Interval, Rules, days_between


def build_intervals(issue: FlowIssue, as_of: dt.datetime, rules: Rules) -> tuple[Interval, ...]:
    """One interval per status visit, oldest first. The last one is open (left = None) and runs to `as_of`.

    The first status is the `from` of the first change (the issue was created there); with no change it is the
    current status. Changes are taken in date order, whatever their order in the extract.
    """
    changes = sorted(issue.changes, key=lambda c: (c.at, c.id))
    status = changes[0].from_status if changes else issue.status
    entered = issue.created
    out: list[Interval] = []
    for change in changes:
        out.append(_interval(issue, len(out) + 1, status, entered, change.at, change.id, as_of, rules))
        status, entered = change.to_status, change.at
    out.append(_interval(issue, len(out) + 1, status, entered, None, None, as_of, rules))
    return tuple(out)


def _interval(issue, seq, status, entered, left, change_id, as_of, rules) -> Interval:
    return Interval(
        key=issue.key, seq=seq, status=status, category=rules.rule(issue.workflow, status).category,
        entered=entered, left=left, duration_days=days_between(entered, left if left is not None else as_of),
        is_open=left is None, change_id=change_id)
