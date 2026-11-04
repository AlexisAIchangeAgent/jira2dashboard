"""Flow Load: work in process per day (docs section 2.2, parameter FlowLoadMode in the Config sheet)."""
import datetime as dt

from jira2dashboard.flow.itemflow import ItemFlow
from jira2dashboard.flow.model import ACTIVE, MODES, UNCLASSIFIED, WAIT

COUNTED = (ACTIVE, WAIT, UNCLASSIFIED)        # categories that are work in process: not Backlog, not terminal
ONE_DAY = dt.timedelta(days=1)


def in_process_at(item: ItemFlow, moment: dt.datetime, mode: str) -> bool:
    """Is the issue work in process at `moment`?

    * Mode A: it has started (first_start <= moment) and sits, at that instant, in an interval
      [entered, left) whose category is not Backlog and not terminal (Done, Removed). Time spent in Backlog after
      a first start is not counted; the time before the return to Backlog is.
    * Mode B: it is counted only from its last StartUtc (after the last return to Backlog) until DoneUtc or
      RemovedUtc, whichever exists; with no StartUtc it is never counted.
    """
    if mode not in MODES:
        raise ValueError(f"FlowLoadMode must be A or B, got {mode!r}")
    if mode == "A":
        if item.first_start is None or item.first_start > moment:
            return False
        return any(i.category in COUNTED and i.entered <= moment and (i.left is None or moment < i.left)
                   for i in item.intervals)
    if item.start is None or item.start > moment:
        return False
    end = item.done if item.done is not None else item.removed
    return end is None or moment < end


def snapshot_days(first_day: dt.datetime, end: dt.datetime, as_of: dt.datetime) -> list[dt.datetime]:
    """The days D of [first_day, end) whose end-of-day snapshot (midnight UTC of D + 1) is not after `as_of`.

    A day that has not finished at `as_of` has no snapshot (open intervals only run to `as_of`).
    """
    days, day = [], dt.datetime(first_day.year, first_day.month, first_day.day)
    while day < end and day + ONE_DAY <= as_of:
        days.append(day)
        day += ONE_DAY
    return days


def daily_load(items: list[ItemFlow], mode: str, team: str, first_day: dt.datetime, end: dt.datetime,
               as_of: dt.datetime) -> dict[dt.datetime, int]:
    """DailyLoad of a team: for each day D, the number of its issues in process at the end of D (UTC)."""
    mine = [i for i in items if i.team == team]
    return {day: sum(1 for i in mine if in_process_at(i, day + ONE_DAY, mode))
            for day in snapshot_days(first_day, end, as_of)}
