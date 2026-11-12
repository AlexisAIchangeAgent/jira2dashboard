"""The six flow metrics per iteration and team (docs/safe-flow-metrics-and-data-checks.md section 2.2), and Feature
Predictability per PI (the variant of Flow Predictability that AL chose)."""
import datetime as dt
from dataclasses import dataclass

from jira2dashboard.flow.checks import CheckSettings, valid_pi
from jira2dashboard.flow.itemflow import ItemFlow
from jira2dashboard.flow.load import ONE_DAY, in_process_at, snapshot_days
from jira2dashboard.flow.model import NOT_COMPUTABLE, UNMAPPED, WORK_TYPES, FlowIssue, Iteration

BUCKETS = WORK_TYPES + (UNMAPPED,)


@dataclass(frozen=True)
class TeamIterationMetrics:
    """The metrics of one team over one iteration. A figure that cannot be computed is NOT_COMPUTABLE."""
    iteration: str
    team: str
    velocity: int                          # issues done in the iteration
    flow_time_count: int                   # issues behind the Flow Time and Flow Efficiency figures
    flow_time_excluded: int                # done in the iteration but left out: no valid start, or reopened
    flow_time_median: float | str
    flow_time_p85: float | str
    flow_load_average: float | str
    flow_load_end: int | str
    flow_efficiency: float | str
    distribution_done: dict[str, float] | str      # work type -> share of the issues done in the iteration
    distribution_wip: dict[str, float] | str       # work type -> share of the work in process at the iteration end
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


def team_iteration_metrics(items: list[ItemFlow], team: str, iteration: Iteration, mode: str,
                        as_of: dt.datetime) -> TeamIterationMetrics:
    mine = [i for i in items if i.team == team]
    done = [i for i in mine if iteration.contains(i.done)]
    counted = [i for i in done if not i.excluded_from_flow_time]

    times = [i.flow_time_days for i in counted]
    median = percentile_inc(times, 0.5) if times else NOT_COMPUTABLE
    p85 = percentile_inc(times, 0.85) if times else NOT_COMPUTABLE

    active = sum(i.active_days for i in counted)
    wait = sum(i.wait_days for i in counted)
    unclassified = sum(i.unclassified_days for i in counted)
    efficiency = active / (active + wait) if counted and unclassified == 0 and active + wait > 0 else NOT_COMPUTABLE

    days = snapshot_days(iteration.start, iteration.end, as_of)
    loads = [sum(1 for i in mine if in_process_at(i, day + ONE_DAY, mode)) for day in days]
    wip_at_end = [i for i in mine if days and in_process_at(i, days[-1] + ONE_DAY, mode)]
    return TeamIterationMetrics(
        iteration=iteration.name, team=team, velocity=len(done), flow_time_count=len(counted),
        flow_time_excluded=len(done) - len(counted), flow_time_median=median, flow_time_p85=p85,
        flow_load_average=sum(loads) / len(loads) if loads else NOT_COMPUTABLE,
        flow_load_end=loads[-1] if loads else NOT_COMPUTABLE,
        flow_efficiency=efficiency, distribution_done=_shares(done), distribution_wip=_shares(wip_at_end),
        flow_predictability=NOT_COMPUTABLE)


def flow_metrics(items: list[ItemFlow], iterations: list[Iteration], mode: str,
                 as_of: dt.datetime) -> list[TeamIterationMetrics]:
    """Every (iteration, team) of the issues, iterations in the given order, teams alphabetical."""
    teams = sorted({i.team for i in items})
    return [team_iteration_metrics(items, team, iteration, mode, as_of) for iteration in iterations for team in teams]


@dataclass(frozen=True)
class PIPredictability:
    """Feature Predictability of one PI, for the ART: the Features done before the end of the PI over those planned in it."""
    pi: str
    planned: int                           # Features whose PI field is this PI, whatever their status
    done: int | None                       # among them, done (DoneUtc) before the end of the PI; None: no end date
    ratio: float | str                     # done / planned; NOT_COMPUTABLE when nothing is planned or the end is unknown


def _pi_order(pi: str) -> tuple[str, str]:
    return pi[6:], pi[3:5]                 # PI#XX-YYYY: by year, then number


def feature_predictability(issues: list[FlowIssue], items: list[ItemFlow], settings: CheckSettings) -> list[PIPredictability]:
    """The variant of Flow Predictability chosen by AL (2026-11-12): Flow Predictability itself needs the business value of the
    PI objectives, which Jira does not hold (docs section 2.2). One row per PI of the sheet PIs, in its order. Planned = a
    Feature (type in FeatureTypes) whose PI field is that PI; a cancelled one stays planned, not done. Done = its DoneUtc is
    before the end of the PI (the end is excluded, as for an iteration): a Feature done later is not done for that PI.
    Then one row per PI that Features carry but the sheet lacks (AL, 2026-11-12), by year then number: without an end date
    it is not computable, but the RTE sees it. Only a valid PI (PI#XX-YYYY) counts: a malformed one is for DC17 and DC29."""
    done_at = {item.key: item.done for item in items}
    features = [i for i in issues if i.type in settings.feature_types]

    def planned_in(pi: str) -> list[str]:
        return [i.key for i in features if (i.pi or "").strip() == pi]

    out = []
    for pi, (_start, end) in settings.pis.items():
        planned = planned_in(pi)
        done = [key for key in planned if done_at.get(key) is not None and done_at[key] < end]
        out.append(PIPredictability(pi, len(planned), len(done), len(done) / len(planned) if planned else NOT_COMPUTABLE))
    in_sheet = {pi.upper() for pi in settings.pis}           # MATCH and COUNTIFS of the workbook ignore the case
    missing = {i.pi for i in features if valid_pi(i.pi) and i.pi.upper() not in in_sheet}
    out += [PIPredictability(pi, len(planned_in(pi)), None, NOT_COMPUTABLE) for pi in sorted(missing, key=_pi_order)]
    return out
