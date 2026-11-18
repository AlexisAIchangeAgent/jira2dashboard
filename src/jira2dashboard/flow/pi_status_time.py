"""PIStatusTime: where the cycle time of the Features goes, status by status, PI by PI (AL, 2026-11-17).

The purpose is to see which statuses take most of the cycle time. For each PI of the sheet PIs, the items are the Features released
in it (DoneUtc in [Start, End) of the PI, flow-time rules: not excluded, so with a valid start); for each in-process status of the
Feature workflow (category Active or Wait, in the order of the sheet Workflows) the time the items spent in it between their StartUtc
(included) and their DoneUtc (excluded), the same window as ItemFlow.ActiveDays / WaitDays: so for a PI the totals of the statuses
add up to the Active + Wait days of those items. AverageDays = TotalDays / Items (an item that never visited the status counts 0);
Share = TotalDays / the total of all the statuses of the PI (the bar of the PI adds up to 100 %). A PI without a released Feature, or
that the sheet PIs does not hold, is not computable. Excel: the table PIStatusTime (formulas.PISTATUSTIME_COLUMNS).
"""
import datetime as dt
from dataclasses import dataclass

from jira2dashboard.flow.itemflow import ItemFlow
from jira2dashboard.flow.model import ACTIVE, ART, NOT_COMPUTABLE, WAIT

FEATURE_WORKFLOW = "Feature"
NO_DEPTH = 1_000_000
IN_CYCLE = (ACTIVE, WAIT)                    # Backlog is not cycle time, Done ends it; Unclassified is not a status anyone chose


@dataclass(frozen=True)
class PIStatusTime:
    pi: str
    status: str
    items: int | str                         # Features released in the PI, or NOT_COMPUTABLE (the PI has no dates)
    total_days: float | str
    average_days: float | str
    share: float | str


def feature_statuses(items: list[ItemFlow], depths: dict[str, int] | None = None) -> list[str]:
    """The statuses of the stack: those the Features (the ART) visited and whose category is Active or Wait (the Jira workflows hold many more, and
    the sheet Workflows only exceptions). In the order of the normal flow: by depth in the transitions of the Feature workflow (`depths`, status
    -> moves from the creation status, flow_depths; a status without depth comes last), then by the earliest position (Seq) at which a Feature
    visited it, then by name (binary). VBA twin: modFlowBuild.StatusNames."""
    first: dict[str, int] = {}
    for item in items:
        if item.team != ART:
            continue
        for interval in item.intervals:
            if interval.category in IN_CYCLE:
                first[interval.status] = min(first.get(interval.status, interval.seq), interval.seq)
    return sorted(first, key=lambda status: (depths.get(status, NO_DEPTH) if depths else NO_DEPTH, first[status], status))


def released_in(items: list[ItemFlow], window: tuple[dt.datetime, dt.datetime]) -> list[ItemFlow]:
    """The Features done in the window, that flow time counts (ItemFlow.excluded_from_flow_time leaves out the others)."""
    start, end = window
    return [i for i in items if i.team == ART and i.done is not None and start <= i.done < end and not i.excluded_from_flow_time]


def days_in(item: ItemFlow, status: str) -> float:
    """The days the item spent in `status` from its StartUtc (included) to its DoneUtc (excluded), in the categories of the cycle."""
    return sum(i.duration_days for i in item.intervals
               if i.status == status and i.category in IN_CYCLE and item.start <= i.entered < item.done)


def pi_status_time(items: list[ItemFlow], statuses: list[str], pi_names: list[str],
                   windows: dict[str, tuple[dt.datetime, dt.datetime]]) -> list[PIStatusTime]:
    """One row per PI (in the order of `pi_names`) and status. A PI that `windows` does not hold has no dates: not computable."""
    out = []
    for pi in pi_names:
        if pi not in windows:
            out += [PIStatusTime(pi, status, NOT_COMPUTABLE, NOT_COMPUTABLE, NOT_COMPUTABLE, NOT_COMPUTABLE) for status in statuses]
            continue
        mine = released_in(items, windows[pi])
        totals = {status: sum(days_in(i, status) for i in mine) for status in statuses}
        whole = sum(totals.values())
        for status in statuses:
            out.append(PIStatusTime(pi, status, len(mine), totals[status],
                                    totals[status] / len(mine) if mine else NOT_COMPUTABLE,
                                    totals[status] / whole if mine and whole > 0 else NOT_COMPUTABLE))
    return out
