"""The six flow metrics per iteration and team (docs/safe-flow-metrics-and-data-checks.md section 2.2)."""
import datetime as dt
from dataclasses import dataclass

from jira2dashboard.flow.itemflow import ItemFlow
from jira2dashboard.flow.load import ONE_DAY, in_process_at, snapshot_days
from jira2dashboard.flow.model import NOT_COMPUTABLE, UNMAPPED, WORK_TYPES, Period

BUCKETS = WORK_TYPES + (UNMAPPED,)


@dataclass(frozen=True)
class TeamPeriodMetrics:
    """The metrics of one team over one period. A figure that cannot be computed is NOT_COMPUTABLE."""
    period: str
    team: str
    velocity: int                          # issues done in the period
    flow_time_count: int                   # issues behind the Flow Time and Flow Efficiency figures
    flow_time_excluded: int                # done in the period but left out: no valid start, or reopened
    flow_time_median: float | str
    flow_time_p85: float | str
    flow_load_average: float | str
    flow_load_end: int | str
    flow_efficiency: float | str
    distribution_done: dict[str, float] | str      # work type -> share of the issues done in the period
    distribution_wip: dict[str, float] | str       # work type -> share of the work in process at the period end
    flow_predictability: str                       # always NOT_COMPUTABLE: there are no PI objectives in Jira


def percentile_inc(values: list[float], p: float) -> float:
    """Same as Excel PERCENTILE.INC: linear interpolation at position (n - 1) * p of the sorted values."""
    ordered = sorted(values)
    position = (len(ordered) - 1) * p
    low = int(position)
    high = min(low + 1, len(ordered) - 1)
    return ordered[low] + (ordered[high] - ordered[low]) * (position - low)


def _bucket(work_type: str | None) -> str:
    return work_type if work_type in WORK_TYPES else UNMAPPED


def _shares(items: list[ItemFlow]):
    if not items:
        return NOT_COMPUTABLE
    counts = {bucket: 0 for bucket in BUCKETS}
    for item in items:
        counts[_bucket(item.work_type)] += 1
    return {bucket: counts[bucket] / len(items) for bucket in BUCKETS}


def team_period_metrics(items: list[ItemFlow], team: str, period: Period, mode: str,
                        as_of: dt.datetime) -> TeamPeriodMetrics:
    mine = [i for i in items if i.team == team]
    done = [i for i in mine if period.contains(i.done)]
    counted = [i for i in done if not i.excluded_from_flow_time]

    times = [i.flow_time_days for i in counted]
    median = percentile_inc(times, 0.5) if times else NOT_COMPUTABLE
    p85 = percentile_inc(times, 0.85) if times else NOT_COMPUTABLE

    active = sum(i.active_days for i in counted)
    wait = sum(i.wait_days for i in counted)
    unclassified = sum(i.unclassified_days for i in counted)
    efficiency = active / (active + wait) if counted and unclassified == 0 and active + wait > 0 else NOT_COMPUTABLE

    days = snapshot_days(period.start, period.end, as_of)
    loads = [sum(1 for i in mine if in_process_at(i, day + ONE_DAY, mode)) for day in days]
    wip_at_end = [i for i in mine if days and in_process_at(i, days[-1] + ONE_DAY, mode)]
    return TeamPeriodMetrics(
        period=period.name, team=team, velocity=len(done), flow_time_count=len(counted),
        flow_time_excluded=len(done) - len(counted), flow_time_median=median, flow_time_p85=p85,
        flow_load_average=sum(loads) / len(loads) if loads else NOT_COMPUTABLE,
        flow_load_end=loads[-1] if loads else NOT_COMPUTABLE,
        flow_efficiency=efficiency, distribution_done=_shares(done), distribution_wip=_shares(wip_at_end),
        flow_predictability=NOT_COMPUTABLE)


def flow_metrics(items: list[ItemFlow], periods: list[Period], mode: str,
                 as_of: dt.datetime) -> list[TeamPeriodMetrics]:
    """Every (period, team) of the issues, periods in the given order, teams alphabetical."""
    teams = sorted({i.team for i in items})
    return [team_period_metrics(items, team, period, mode, as_of) for period in periods for team in teams]
