"""The PI field of an issue through time, from the changelog: which PI an issue carried at any moment.

The current value of the PI field says where an issue is now, not where it was engaged: a Feature deferred out of its PI has its
PI field emptied or moved, and a Feature added late has it from the day it was added. Feature Predictability measures what was
engaged, so it reads the history. Without any change in the changelog the current value holds for the whole life of the issue.
Pure; the workbook builds the same intervals in the table PIIntervals (modFlow).
"""
import datetime as dt
from dataclasses import dataclass

from jira2dashboard.flow.model import FlowIssue, change_order


@dataclass(frozen=True)
class PIInterval:
    """The issue carried `pi` from `start` (included) to `end` (excluded); end None = still."""
    key: str
    pi: str
    start: dt.datetime
    end: dt.datetime | None


def _clean(value: str | None) -> str | None:
    """The PI as typed, without the blanks around it; an empty field has no PI."""
    text = (value or "").strip()
    return text or None


def pi_intervals(issue: FlowIssue) -> list[PIInterval]:
    """Every stretch of time in which the issue carried a PI, oldest first (a stretch with no PI is not listed). The first
    value is what the first change left (`from_value`), else the current value: it holds from the creation of the issue."""
    changes = sorted(issue.pi_changes, key=change_order)
    current = _clean(changes[0].from_value if changes else issue.pi)
    since, out = issue.created, []
    for change in changes:
        if current and change.at > since:
            out.append(PIInterval(issue.key, current, since, change.at))
        current, since = _clean(change.to_value), max(change.at, since)
    if current:
        out.append(PIInterval(issue.key, current, since, None))
    return _merged(out)


def _merged(intervals: list[PIInterval]) -> list[PIInterval]:
    """One stretch where the same PI is carried without a break (a change that writes the PI it already had splits nothing): the
    workbook counts the stretches, so a Feature must not have two touching ones."""
    out: list[PIInterval] = []
    for interval in intervals:
        if out and out[-1].pi == interval.pi and out[-1].end == interval.start:
            out[-1] = PIInterval(interval.key, interval.pi, out[-1].start, interval.end)
        else:
            out.append(interval)
    return out


def interval_at(intervals: list[PIInterval], moment: dt.datetime) -> PIInterval | None:
    """The stretch that holds `moment` (the end of a stretch is excluded, its start included); None when none, or before creation."""
    for interval in intervals:
        if interval.start <= moment and (interval.end is None or moment < interval.end):
            return interval
    return None


def pi_at(intervals: list[PIInterval], moment: dt.datetime) -> str | None:
    """The PI carried at `moment`; None when none, or before creation."""
    found = interval_at(intervals, moment)
    return found.pi if found else None
