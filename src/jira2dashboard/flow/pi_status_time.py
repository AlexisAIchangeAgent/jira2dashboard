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
from jira2dashboard.params.model import ParamSheet

FEATURE_WORKFLOW = "Feature"
IN_CYCLE = (ACTIVE, WAIT)                    # Backlog is not cycle time, Done ends it; Unclassified is not a status anyone chose


@dataclass(frozen=True)
class PIStatusTime:
    pi: str
    status: str
    items: int | str                         # Features released in the PI, or NOT_COMPUTABLE (the PI has no dates)
    total_days: float | str
    average_days: float | str
    share: float | str


def feature_statuses(sheets: list[ParamSheet] | tuple[ParamSheet, ...]) -> list[str]:
    """The statuses of the stack: the Feature workflow's Active and Wait statuses, in the ROW order of the sheet Workflows (the column
    Order is empty for a free status, and VBA and Python would not break that tie alike)."""
    workflows = next(s for s in sheets if s.name == "Workflows")
    return [row[2] for row in workflows.rows if row[0] == FEATURE_WORKFLOW and row[3] in IN_CYCLE]


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
