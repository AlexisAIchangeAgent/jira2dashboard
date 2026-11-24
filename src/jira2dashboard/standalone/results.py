"""The result sheets of the Python path, computed by the flow reference, with the column names of the workbook's tables.

ItemFlow (the columns the reference has, then a 1/0 column per data check), DailyLoad (every day x team of the
iterations, empty after the as-of date), Metrics (every iteration x team of the sheet Projects, as the workbook), FillRates
and DataChecks. Same names as formulas.TABLES, so whatever reads the workbook's tables (the Confluence page of the RTE)
reads these too.
"""
import datetime as dt
from dataclasses import dataclass, field, replace

from jira2dashboard.flow import dashboard_spec as D
from jira2dashboard.flow import formulas as F
from jira2dashboard.flow.blueprint import team_names
from jira2dashboard.flow.checks import CHECKS, PER_ISSUE, FillRate, fill_rates, run_checks
from jira2dashboard.flow.itemflow import ItemFlow, item_flow
from jira2dashboard.flow.iteration_features import IterationFeatures, iteration_features
from jira2dashboard.flow.load import ONE_DAY, daily_load, snapshot_days
from jira2dashboard.flow.pi_history import pi_intervals
from jira2dashboard.flow.reconciliation import reconcile
from jira2dashboard.flow.trace import MetricItems, metric_items, trace
from jira2dashboard.flow.pi_status_time import PIStatusTime, feature_statuses, pi_status_time
from jira2dashboard.flow.metrics import (BUCKETS, PIPredictability, TeamIterationMetrics, pi_predictability,
                                         team_iteration_metrics)
from jira2dashboard.flow.model import ART, NOT_COMPUTABLE, OTHER_TYPES, UNMAPPED, WORK_TYPES, FlowIssue, work_type_bucket
from jira2dashboard.flow.metric_checks import (FlowDistCheck, in_process_at_end, FlowEfficiencyCheck, FlowLoadCheck, FlowTimeCheck, flow_distribution_check,
                                               flow_efficiency_check, flow_load_check, flow_time_check)
from jira2dashboard.flow.velocity_check import VelocityCheck, velocity_check
from jira2dashboard.standalone.params_book import RunSettings

ITEMFLOW = ("Key", "Project", "Type", "WorkType", "Role", "Workflow", "Team", "CreatedUtc", "FirstStartUtc", "StartUtc",
            "BacklogReturns", "DoneUtc", "RemovedUtc", "FlowTimeDays", "LeadTimeDays", "ActiveDays", "WaitDays",
            "UnclassifiedDays", "Reopened", "SkippedStart", "Excluded", "WorkTypeBucket", "Parent", "StatusCategory",
            "ResolvedUtc", "FixVersions", "StoryPoints", "PI", "BV", "TeamValue", "SprintState", "SprintStart",
            "SprintEnd", "SprintCount", "TargetStart", "TargetEnd", "KeyChanges") + PER_ISSUE + (
            "DoneIteration", "DoneIterationStart", "DoneIterationEnd", "DonePI", "Account", "WhyNotInFlowTime", "BoundaryHours", "JiraWorkflow")
DAILY_LOAD = F.TABLES["DailyLoad"]
METRICS = F.TABLES["Metrics"]
PI_METRICS = F.TABLES["PIMetrics"]
PI_STATUS_TIME = F.TABLES["PIStatusTime"]
VELOCITY_CHECK = F.TABLES["VelocityCheck"]
FLOW_TIME_CHECK = F.TABLES["FlowTimeCheck"]
FLOW_EFFICIENCY_CHECK = F.TABLES["FlowEfficiencyCheck"]
FLOW_LOAD_CHECK = F.TABLES["FlowLoadCheck"]
FLOW_DISTRIBUTION_CHECK = F.TABLES["FlowDistCheck"]
RECONCILIATION = tuple(name for name, _ in F.RECONCILIATION_COLUMNS)
ITERATION_FEATURES = F.ITERATIONFEATURES_KEYS + F.ITERATIONFEATURES_NAMES
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
    iteration_features: list[IterationFeatures]   # the Features per iteration, PI and team: what the PI figures are summed from
    status_time: list[PIStatusTime]           # where the cycle time of the Features goes, status by status, PI by PI
    velocity_check: list[VelocityCheck]       # the Velocity counted again from the status changes, every iteration x team
    flow_time_check: list[FlowTimeCheck] = field(default_factory=list)      # why the Flow Time is what it is, every iteration x team
    flow_efficiency_check: list[FlowEfficiencyCheck] = field(default_factory=list)      # why the Flow Efficiency is what it is
    flow_load_check: list[FlowLoadCheck] = field(default_factory=list)      # why the Flow Load is what it is
    flow_distribution_check: list[FlowDistCheck] = field(default_factory=list)      # what the Flow Distribution counts


def iteration_span(iterations) -> tuple[dt.datetime, dt.datetime]:
    """The first start and the last end of the iterations, whatever their order in the sheet (as modFlowBuild.DaysOf)."""
    return min(i.start for i in iterations), max(i.end for i in iterations)


def fill_teams(issues: list[FlowIssue]) -> list[str]:
    """The teams of FillRates: those of the issues but "ART (other types)" (the workbook lists ART and the Team projects)."""
    return sorted({i.team for i in issues} - {OTHER_TYPES})


def with_jira_workflow(issues: list[FlowIssue], s: RunSettings) -> list[FlowIssue]:
    """Each issue with the Jira workflow of its project (the prefix of its key) and issue type, as the column ItemFlow.JiraWorkflow reads JiraIssueTypes."""
    return [replace(i, jira_workflow=s.workflows.workflow_of.get((i.key.rsplit("-", 1)[0], i.type), "")) for i in issues]


def team_workflow(team: str) -> str:
    """The workflow kind of a team of the metrics: the ART follows the Feature workflow, a team of a Team project the Team workflow."""
    return "Feature" if team == ART else "Team"


def compute(issues: list[FlowIssue], s: RunSettings) -> Results:
    issues = with_jira_workflow(issues, s)
    items = [item_flow(i, s.as_of, s.rules) for i in issues]
    teams = team_names(s.projects)
    first, end = iteration_span(s.iterations)
    flags = run_checks(issues, s.as_of, s.rules, s.moves.with_jira(s.workflows), s.checks)
    fill = fill_rates(issues, s.checks, fill_teams(issues), sorted({i.type for i in issues}))
    rows = iteration_features(issues, items, s.iterations, s.checks.pis, s.checks.feature_types)
    predictability = pi_predictability(rows, issues, s.checks, s.predictability_mode)
    pi_names = [p.pi for p in predictability] or [D.NO_PI]               # the keys of PIMetrics, also those of PIStatusTime
    status_time = pi_status_time(items, feature_statuses(items, s.workflows.feature_depths([k for k, role, _t in s.projects if role == "Program"], s.checks.feature_types)) or [D.NO_STATUS], pi_names, s.checks.pis)
    metrics = [team_iteration_metrics(items, team, iteration, s.mode, s.as_of, s.rules.waits, s.rules.has_start(team_workflow(team)))
               for iteration in s.iterations for team in teams]
    iteration_of = {i.name: i for i in s.iterations}
    return Results(items, flags, teams, metrics,
                   {team: daily_load(items, s.mode, team, first, end, s.as_of, s.rules.has_start(team_workflow(team))) for team in teams}, fill,
                   predictability, rows, status_time,
                   [velocity_check(issues, items, iteration, team, s.rules, s.checks.feature_types)
                    for iteration in s.iterations for team in teams],
                   [flow_time_check(issues, items, iteration_of[m.iteration], m.team, s.rules, m) for m in metrics],
                   [flow_efficiency_check(items, iteration_of[m.iteration], m.team, s.rules, m) for m in metrics],
                   [flow_load_check(items, iteration_of[m.iteration], m.team, s.rules, s.mode, s.as_of, m) for m in metrics],
                   [flow_distribution_check(items, iteration_of[m.iteration], m.team, m, s.work_type_field) for m in metrics])


def _bucket(work_type) -> str:
    return work_type_bucket(work_type)


def itemflow_rows(results: Results, rows: list[tuple], issues: list[FlowIssue],
                  projects: list[tuple[str, str, str]], iterations=(), pis=None) -> list[tuple]:
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
                   + tuple(1 if check in ids else 0 for check in PER_ISSUE)
                   + _trace_cells(item, list(iterations), pis or {}) + (issue.jira_workflow,))
    return out


def _trace_cells(item: ItemFlow, iterations: list, pis: dict) -> tuple:
    """DoneIteration .. BoundaryHours (flow/trace.py): "" or None where the workbook's formula gives the empty text."""
    t = trace(item, iterations, pis)
    return (t.done_iteration, t.start, t.end, t.done_pi, t.account, t.why_not_in_flow_time, t.boundary_hours)


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


def iterationfeatures_rows(results: Results) -> list[tuple]:
    """The rows of IterationFeatures (flow/iteration_features.py), in the column order of the table."""
    return [(r.iteration, r.start, r.end, r.team, r.pi or None, r.baseline, r.engaged, r.added, r.deferred, r.released,
             r.released_planned, r.released_added, r.carry_over) for r in results.iteration_features]


def pistatustime_rows(results: Results, s: RunSettings) -> list[tuple]:
    """The rows of PIStatusTime (formulas.PISTATUSTIME_COLUMNS), one per PI of PIMetrics and Active / Wait status of the Feature workflow."""
    return [(r.pi, r.status, *s.checks.pis.get(r.pi, (None, None)), r.items, r.total_days, r.average_days, r.share)
            for r in results.status_time]


def reconciliation_rows(results: Results, issues: list[FlowIssue], s: RunSettings) -> list[tuple]:
    """The rows of the table Reconciliation (formulas.RECONCILIATION_COLUMNS): control, assertion, expected, found, difference, result."""
    return [(c.name, c.assertion, c.expected, c.found, c.found - c.expected, c.verdict)
            for c in reconcile(issues, results.items, results.metrics, list(s.iterations), s.checks.pis, s.rules, results.flags, mode=s.mode,
                               as_of=s.as_of, features=results.iteration_features)]


def metricitems_sheet(results: Results, s: RunSettings) -> tuple[tuple, list[tuple]]:
    """The sheet MetricItems of the Python path: the header (the names of the three blocks with an empty name between them) and the rows
    (the three lists of flow/trace.metric_items side by side, the shorter ones padded with empty cells)."""
    iterations = {p.name: p for p in s.iterations}
    lists = metric_items(results.items, list(s.iterations), s.checks.pis)
    at_end = in_process_at_end(results.items, [(iterations[m.iteration], m.team) for m in results.metrics], s.rules, s.mode, s.as_of)
    return metricitems_grid(replace(lists, at_end=at_end))


def metricitems_grid(lists: MetricItems) -> tuple[tuple, list[tuple]]:
    blocks = [[name for name, _ in block] for block in F.METRICITEMS_BLOCKS]
    header = tuple(name for names in blocks for name in names + [""])[:-1]
    data = [lists.done, lists.left_out, lists.cutoff, lists.at_end]
    rows = []
    for n in range(max(len(d) for d in data)):
        row: list = []
        for names, block in zip(blocks, data):
            row += list(block[n]) if n < len(block) else [None] * len(names)
            row.append(None)
        rows.append(tuple(row[:-1]))
    return header, rows


def velocitycheck_rows(results: Results, s: RunSettings) -> list[tuple]:
    """The rows of VelocityCheck (formulas.VELOCITYCHECK_COLUMNS), one per iteration and team, in the order of Metrics."""
    iterations = {p.name: p for p in s.iterations}
    return [(r.iteration, iterations[r.iteration].start, iterations[r.iteration].end, r.team, r.velocity, r.velocity_features,
             r.other_types, r.created_done, r.changed_to_end, r.changed_to_end_features, r.reopened_later, r.gap, r.reason)
            for r in results.velocity_check]


def flowtimecheck_rows(results: Results, s: RunSettings) -> list[tuple]:
    """The rows of FlowTimeCheck (formulas.FLOWTIMECHECK_COLUMNS), one per iteration and team, in the order of Metrics."""
    iterations = {p.name: p for p in s.iterations}
    return [(r.iteration, iterations[r.iteration].start, iterations[r.iteration].end, r.team, r.done, r.excl_reopened, r.excl_no_start_rule,
             r.excl_never_started, r.excl_back_to_backlog, r.counted, r.gap, r.reason, r.cause) for r in results.flow_time_check]


def flowefficiencycheck_rows(results: Results, s: RunSettings) -> list[tuple]:
    """The rows of FlowEfficiencyCheck (formulas.FLOWEFFICIENCYCHECK_COLUMNS), one per iteration and team, in the order of Metrics."""
    iterations = {p.name: p for p in s.iterations}
    return [(r.iteration, iterations[r.iteration].start, iterations[r.iteration].end, r.team, r.counted, r.with_unclassified, r.active_days,
             r.wait_days, r.wait_statuses, "" if r.efficiency is None else r.efficiency, r.gap, r.reason, r.cause)
            for r in results.flow_efficiency_check]


def flowloadcheck_rows(results: Results, s: RunSettings) -> list[tuple]:
    """The rows of FlowLoadCheck (formulas.FLOWLOADCHECK_COLUMNS), one per iteration and team, in the order of Metrics."""
    iterations = {p.name: p for p in s.iterations}
    return [(r.iteration, iterations[r.iteration].start, iterations[r.iteration].end, r.team, r.days, r.days_measured, r.start_statuses,
             "" if r.average is None else r.average, "" if r.at_end is None else r.at_end, r.gap, r.reason) for r in results.flow_load_check]


def flowdistributioncheck_rows(results: Results, s: RunSettings) -> list[tuple]:
    """The rows of FlowDistCheck (formulas.FLOWDISTRIBUTIONCHECK_COLUMNS), one per iteration and team, in the order of Metrics."""
    iterations = {p.name: p for p in s.iterations}
    return [(r.iteration, iterations[r.iteration].start, iterations[r.iteration].end, r.team, r.done, *r.counts, r.gap, r.reason, r.note)
            for r in results.flow_distribution_check]


def piinterval_rows(issues: list[FlowIssue]) -> list[tuple]:
    """The rows of PIIntervals (formulas.PIINTERVALS_RAW): every stretch in which an issue carried a PI, from its history."""
    return [(i.key, i.pi, i.start, i.end, i.end is None) for issue in issues for i in pi_intervals(issue)]


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
