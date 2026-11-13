"""The result sheets of the Python path, computed by the flow reference, with the column names of the workbook's tables.

ItemFlow (the columns the reference has, then a 1/0 column per data check), DailyLoad (every day x team of the
iterations, empty after the as-of date), Metrics (every iteration x team of the sheet Projects, as the workbook), FillRates
and DataChecks. Same names as formulas.TABLES, so whatever reads the workbook's tables (the Confluence page of the RTE)
reads these too.
"""
import datetime as dt
from dataclasses import dataclass

from jira2dashboard.flow import dashboard_spec as D
from jira2dashboard.flow import formulas as F
from jira2dashboard.flow.blueprint import team_names
from jira2dashboard.flow.checks import CHECKS, PER_ISSUE, FillRate, fill_rates, run_checks
from jira2dashboard.flow.itemflow import ItemFlow, item_flow
from jira2dashboard.flow.load import ONE_DAY, daily_load, snapshot_days
from jira2dashboard.flow.metrics import (BUCKETS, PIPredictability, TeamIterationMetrics, feature_predictability,
                                         team_iteration_metrics)
from jira2dashboard.flow.model import NOT_COMPUTABLE, UNMAPPED, WORK_TYPES, FlowIssue
from jira2dashboard.standalone.params_book import RunSettings

ITEMFLOW = ("Key", "Project", "Type", "WorkType", "Role", "Workflow", "Team", "CreatedUtc", "FirstStartUtc", "StartUtc",
            "BacklogReturns", "DoneUtc", "RemovedUtc", "FlowTimeDays", "LeadTimeDays", "ActiveDays", "WaitDays",
            "UnclassifiedDays", "Reopened", "SkippedStart", "Excluded", "WorkTypeBucket", "Parent", "StatusCategory",
            "ResolvedUtc", "FixVersions", "StoryPoints", "PI", "BV", "TeamValue", "SprintState", "SprintStart",
            "SprintEnd", "SprintCount", "TargetStart", "TargetEnd", "KeyChanges") + PER_ISSUE
DAILY_LOAD = F.TABLES["DailyLoad"]
METRICS = F.TABLES["Metrics"]
PI_METRICS = F.TABLES["PIMetrics"]
FILL_RATES = F.TABLES["FillRates"]
DATA_CHECKS = F.TABLES["DataChecks"]
_RAW = {name: n for n, name in enumerate(F.ISSUES)}


@dataclass
class Results:
    items: list[ItemFlow]
    flags: dict[str, frozenset[str]]          # key -> ids of the checks it fails
    teams: list[str]                          # ART, then the teams of the Team projects (sheet Projects)
    metrics: list[TeamIterationMetrics]          # every iteration x team
    loads: dict[str, dict[dt.datetime, int]]  # team -> day -> items in process at the end of the day
    fill: list[FillRate]
    predictability: list[PIPredictability]   # Feature Predictability per PI of the sheet PIs (ART)


def iteration_span(iterations) -> tuple[dt.datetime, dt.datetime]:
    """The first start and the last end of the iterations, whatever their order in the sheet (as modFlowBuild.DaysOf)."""
    return min(i.start for i in iterations), max(i.end for i in iterations)


def compute(issues: list[FlowIssue], s: RunSettings) -> Results:
    items = [item_flow(i, s.as_of, s.rules) for i in issues]
    teams = team_names(s.projects)
    first, end = iteration_span(s.iterations)
    flags = run_checks(issues, s.as_of, s.rules, s.moves, s.checks)
    fill = fill_rates(issues, s.checks, sorted({i.team for i in issues}), sorted({i.type for i in issues}))
    return Results(items, flags, teams,
                   [team_iteration_metrics(items, team, iteration, s.mode, s.as_of) for iteration in s.iterations for team in teams],
                   {team: daily_load(items, s.mode, team, first, end, s.as_of) for team in teams}, fill,
                   feature_predictability(issues, items, s.checks))


def _bucket(work_type) -> str:
    return work_type if work_type in WORK_TYPES else UNMAPPED


def itemflow_rows(results: Results, rows: list[tuple], issues: list[FlowIssue],
                  projects: list[tuple[str, str, str]]) -> list[tuple]:
    roles = {key: role for key, role, _team in projects}
    raw = {row[0]: row for row in rows}
    out = []
    for item, issue in zip(results.items, issues):
        row = raw[item.key]
        ids = results.flags.get(item.key, frozenset())
        out.append((item.key, row[_RAW["Project"]], item.type, item.work_type, roles.get(row[_RAW["Project"]], ""),
                    issue.workflow, item.team, item.created, item.first_start, item.start, item.backlog_returns,
                    item.done, item.removed, item.flow_time_days, item.lead_time_days, item.active_days,
                    item.wait_days, item.unclassified_days, item.reopened, "Y" if item.skipped_start else "N",
                    1 if item.excluded_from_flow_time else 0, _bucket(item.work_type))
                   + tuple(row[_RAW[name]] for name in ITEMFLOW[22:37])
                   + tuple(1 if check in ids else 0 for check in PER_ISSUE))
    return out


def days_of(first: dt.datetime, end: dt.datetime) -> list[dt.datetime]:
    """Every day from the first iteration start to the last iteration end (excluded)."""
    days, day = [], dt.datetime(first.year, first.month, first.day)
    while day < end:
        days.append(day)
        day += ONE_DAY
    return days


def dailyload_rows(results: Results, first: dt.datetime, end: dt.datetime) -> list[tuple]:
    """Date | Team | Load for every day of the iterations and every team; Load empty when the day has no snapshot yet."""
    return [(day, team, results.loads[team].get(day)) for day in days_of(first, end) for team in results.teams]


def _figure(value):
    return NOT_COMPUTABLE if value == NOT_COMPUTABLE else value


def pimetrics_rows(results: Results, s: RunSettings) -> list[tuple]:
    """The rows of PIMetrics as the workbook computes them: one per PI of the sheet PIs, then one per PI of a Feature that
    the sheet lacks (no dates), or the one row NO_PI."""
    if not results.predictability:
        return [(D.NO_PI, None, None, 0, None, NOT_COMPUTABLE)]
    return [(p.pi, *s.checks.pis.get(p.pi, (None, None)), p.planned, p.done, p.ratio) for p in results.predictability]


def metrics_rows(results: Results, s: RunSettings) -> list[tuple]:
    iterations = {p.name: p for p in s.iterations}
    out = []
    for m in results.metrics:
        iteration = iterations[m.iteration]
        days = snapshot_days(iteration.start, iteration.end, s.as_of)
        end_day = days[-1] if days else None
        done, wip = m.distribution_done, m.distribution_wip
        load_end = m.flow_load_end
        if end_day is None:
            wip_counts = (None,) * len(BUCKETS)
        elif wip == NOT_COMPUTABLE:
            wip_counts = (0,) * len(BUCKETS)
        else:
            wip_counts = tuple(round(wip[b] * load_end) for b in BUCKETS)
        out.append((m.iteration, iteration.start, iteration.end, m.team, m.velocity, m.flow_time_count, m.flow_time_excluded,
                    _figure(m.flow_time_median), _figure(m.flow_time_p85), _figure(m.flow_load_average), end_day,
                    _figure(load_end), _figure(m.flow_efficiency))
                   + tuple(NOT_COMPUTABLE if done == NOT_COMPUTABLE else done[b] for b in BUCKETS)
                   + wip_counts
                   + tuple(NOT_COMPUTABLE if wip == NOT_COMPUTABLE else wip[b] for b in BUCKETS)
                   + (m.flow_predictability,))
    return out


def fillrates_rows(fill: list[FillRate]) -> list[tuple]:
    return [(r.field, r.team, r.type, r.count, r.filled, r.filled / r.count if r.count else None, 1 if r.below else 0)
            for r in fill]


def data_check_rows(flags: dict[str, frozenset[str]], fill: list[FillRate], settings) -> list[tuple]:
    """Id | Title | Severity | Need1-3 | Source | Ready | Items | Result: the number of items each check flags (the
    fill rates below the threshold for DC25), "not computable" while a field it reads is not configured."""
    below = sum(1 for r in fill if r.below)
    out = []
    for check in CHECKS:
        needs = (check.needs + ("", "", ""))[:3]
        if settings.ready(check):
            count = below if check.source == "FillRates" else sum(1 for ids in flags.values() if check.id in ids)
            tail = (1, count, "OK" if count == 0 else "KO")
        else:
            tail = (0, None, NOT_COMPUTABLE)
        out.append((check.id, check.title, check.severity) + needs + (check.source,) + tail)
    return out
