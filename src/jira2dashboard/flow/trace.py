"""The trace of a ticket: where it sits in the figures (AL, 2026-11-18: "c'est de la comptabilite", every ticket in exactly one account).

For each ticket of ItemFlow:
* DoneIteration (and its Start, End): the iteration whose window [Start, End) holds its DoneUtc, "" when none (or not done); DonePI the same for the PIs
  of the sheet PIs that have dates;
* Account: the ONE bucket of the population it sits in (a partition, so the buckets add up to the tickets read): done in an iteration (it is in
  Velocity, Flow Time, Flow Efficiency and Flow Distribution), done outside the iterations, left out because it is not a Feature (the team
  "ART (other types)"), removed, open;
* WhyNotInFlowTime: "Reopened" and / or "No valid start", the two reasons that leave a done ticket out of Flow Time and Flow Efficiency;
* BoundaryHours: how many hours its DoneUtc is from the nearest iteration boundary (cut-off test: a ticket done a few hours from a boundary is the
  one that could fall on the other side).
Excel: the columns of ItemFlow (formulas.ITEMFLOW_COLUMNS), the lists of the sheet MetricItems and the table Reconciliation are built on them.
"""
import datetime as dt
from dataclasses import dataclass

from jira2dashboard.flow.itemflow import ItemFlow
from jira2dashboard.flow.model import OTHER_TYPES, Iteration

ACCOUNT_COUNTED = "Done in an iteration"
ACCOUNT_OUTSIDE = "Done outside the iterations"
ACCOUNT_OTHER = "Left out: not a Feature"
ACCOUNT_REMOVED = "Removed"
ACCOUNT_OPEN = "Open"
ACCOUNTS = (ACCOUNT_COUNTED, ACCOUNT_OUTSIDE, ACCOUNT_OTHER, ACCOUNT_REMOVED, ACCOUNT_OPEN)
REOPENED, NO_START = "Reopened", "No valid start"


@dataclass(frozen=True)
class Trace:
    done_iteration: str
    start: dt.datetime | None            # of that iteration
    end: dt.datetime | None
    done_pi: str
    account: str
    why_not_in_flow_time: str
    boundary_hours: float | None         # None when the ticket is not done


def done_iteration(item: ItemFlow, iterations: list[Iteration]) -> Iteration | None:
    """The first iteration (in the order of the sheet) that holds the DoneUtc: the iterations of a sheet do not overlap."""
    return next((i for i in iterations if i.contains(item.done)), None)


def account(item: ItemFlow, in_iteration: bool) -> str:
    if item.team == OTHER_TYPES:
        return ACCOUNT_OTHER
    if item.done is not None:
        return ACCOUNT_COUNTED if in_iteration else ACCOUNT_OUTSIDE
    return ACCOUNT_REMOVED if item.removed is not None else ACCOUNT_OPEN


def why_not_in_flow_time(item: ItemFlow) -> str:
    """The reasons of ItemFlow.excluded_from_flow_time, reopened first; "" for a ticket that Flow Time does not leave out."""
    return ", ".join(reason for reason, applies in ((REOPENED, item.reopened > 0), (NO_START, item.skipped_start)) if applies)


def boundary_hours(item: ItemFlow, iterations: list[Iteration]) -> float | None:
    """The distance in hours from DoneUtc to the nearest start or end of an iteration."""
    if item.done is None or not iterations:
        return None
    return min(abs((item.done - boundary).total_seconds()) for i in iterations for boundary in (i.start, i.end)) / 3600


def trace(item: ItemFlow, iterations: list[Iteration], pis: dict[str, tuple[dt.datetime | None, dt.datetime | None]]) -> Trace:
    found = done_iteration(item, iterations)
    pi = next((name for name, (start, end) in pis.items()
               if start is not None and end is not None and item.done is not None and start <= item.done < end), "")
    return Trace(found.name if found else "", found.start if found else None, found.end if found else None, pi,
                 account(item, found is not None), why_not_in_flow_time(item), boundary_hours(item, iterations))


# --- the lists of the sheet MetricItems (the tickets behind the figures) ---------------------------------------------------------

CUTOFF_HOURS = 24                       # a ticket done this close to an iteration boundary is listed for a cut-off review


@dataclass(frozen=True)
class MetricItems:
    """Three lists, as the three spilled blocks of the sheet MetricItems.
    done: iteration, start, end, team, key, type, DoneUtc, FlowTimeCounted (1/0), FlowTimeDays ("" when none), WhyNotInFlowTime - the tickets behind
    Velocity, Flow Time, Flow Efficiency and Flow Distribution (the Velocity of a team and iteration = its rows), by iteration, team, date, key;
    left_out: key, type, team, iteration, DoneUtc, account - the done tickets that no figure counts, with the reason;
    cutoff: key, team, iteration, DoneUtc, hours to the boundary (1 decimal) - done within CUTOFF_HOURS of a boundary, by distance."""
    done: list[tuple]
    left_out: list[tuple]
    cutoff: list[tuple]


def metric_items(items: list[ItemFlow], iterations: list[Iteration],
                 pis: dict[str, tuple[dt.datetime | None, dt.datetime | None]]) -> MetricItems:
    done, left_out, cutoff = [], [], []
    for item in items:
        t = trace(item, iterations, pis)
        if t.account == ACCOUNT_COUNTED:
            done.append((t.done_iteration, t.start, t.end, item.team, item.key, item.type, item.done, 0 if item.excluded_from_flow_time else 1,
                         "" if item.flow_time_days is None else item.flow_time_days, t.why_not_in_flow_time))
        if item.done is not None and t.account != ACCOUNT_COUNTED:
            left_out.append((item.key, item.type, item.team, t.done_iteration, item.done, t.account))
        if t.boundary_hours is not None and t.boundary_hours <= CUTOFF_HOURS and t.account != ACCOUNT_OTHER:
            cutoff.append((item.key, item.team, t.done_iteration, item.done, round(t.boundary_hours, 1)))
    done.sort(key=lambda r: (r[1], r[3], r[6], r[4]))
    left_out.sort(key=lambda r: (r[5], r[4], r[0]))
    cutoff.sort(key=lambda r: (r[4], r[0]))
    return MetricItems(done, left_out, cutoff)
