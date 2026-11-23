"""What the generator writes into the workbook so that VBA can build the flow sheets without Python (pure).

Sheet FlowFormulas (Table | Column | Formula): the columns of every table, in order. An empty Formula marks a column of
values: raw data written by the extraction, or keys written by modFlowBuild. The rows of the tables "Controls" (named
cells) and "Dashboard" (settings) use the third column for a formula or a plain value. Sheet ChartSpec: one row per chart, with
the formula of its points. modFlowBuild reads both after each extraction: formulas and charts are DATA of the workbook,
tested here without Excel and against the Python reference by the scenario flow_excel.feature.
"""
from jira2dashboard.flow import check_formulas as C
from jira2dashboard.flow import dashboard_spec as D
from jira2dashboard.flow import formulas as F
from jira2dashboard.flow import reconciliation as R
from jira2dashboard.flow.checks import CHECKS
from jira2dashboard.flow.model import ART

SHEET_FORMULAS = "FlowFormulas"
SHEET_CHARTS = "ChartSpec"
SHEET_CHECKS = "ChecksSpec"
FORMULAS_HEADER = ("Table", "Column", "Formula")
CHECKS_HEADER = C.DATACHECKS_KEYS
CHARTS_HEADER = ("Name", "Title", "ChartType", "AxisFormat", "SeriesNames", "PointFormula", "NoteFormula",
                 "AnchorRow", "AnchorColumn", "Rows", "Columns", "Axis", "Band", "Sheet", "Labels")

CONTROL_NAMES = ("AsOfUtc", "cfg_FlowLoadMode", "cfg_TransitionsLast", "cfg_PIsLast", "cfg_PredictabilityMode")   # the control cells (sheet Controls, column B)


def formula_rows() -> list[tuple[str, str, str]]:
    rows: list[tuple[str, str, str]] = []
    rows += [("Issues", c, "") for c in F.ISSUES]
    rows += [("StatusIntervals", c, "") for c in F.INTERVALS_RAW]
    rows += [("StatusIntervals", name, text) for name, text in F.INTERVAL_HELPERS]
    rows += [("ItemFlow", name, text) for name, text in F.ITEMFLOW_COLUMNS]
    rows += [("DailyLoad", "Date", ""), ("DailyLoad", "Team", ""), ("DailyLoad", "Load", F.DAILYLOAD_LOAD)]
    rows += [("IssueLinks", key, "") for key in C.LINKS_RAW]
    rows += [("IssueLinks", name, text) for name, text in C.LINKS_COLUMNS]
    rows += [("FillRates", key, "") for key in C.FILL_KEYS]
    rows += [("FillRates", name, text) for name, text in C.FILL_COLUMNS]
    rows += [("FillFields", name, need) for name, need in C.FILL_FIELD_NEEDS]       # field, Config parameter it needs
    rows += [("DataChecks", key, "") for key in C.DATACHECKS_KEYS]
    rows += [("DataChecks", name, text) for name, text in C.DATACHECKS_COLUMNS]
    rows += [("Metrics", key, "") for key in F.METRICS_KEYS]
    rows += [("PIIntervals", key, "") for key in F.PIINTERVALS_RAW]
    rows += [("PIIntervals", name, text) for name, text in F.PIINTERVAL_HELPERS]
    rows += [("IterationFeatures", key, "") for key in F.ITERATIONFEATURES_KEYS]
    rows += [("IterationFeatures", name, text) for name, text in F.ITERATIONFEATURES_COLUMNS]
    rows += [("PIMetrics", key, "") for key in F.PIMETRICS_KEYS]
    rows += [("PIMetrics", name, text) for name, text in F.PIMETRICS_COLUMNS]
    rows += [("PIStatusTime", key, "") for key in F.PISTATUSTIME_KEYS]
    rows += [("PIStatusTime", name, text) for name, text in F.PISTATUSTIME_COLUMNS]
    rows += [("VelocityCheck", key, "") for key in F.VELOCITYCHECK_KEYS]
    rows += [("VelocityCheck", name, text) for name, text in F.VELOCITYCHECK_COLUMNS]
    rows += [("FlowTimeCheck", key, "") for key in F.FLOWTIMECHECK_KEYS]
    rows += [("FlowTimeCheck", name, text) for name, text in F.FLOWTIMECHECK_COLUMNS]
    rows += [("Metrics", name, text) for name, text in F.METRICS_COLUMNS]
    rows += [("Reconciliation", name, text) for name, text in F.RECONCILIATION_COLUMNS]
    rows += [("ReconciliationControl", name, assertion) for name, assertion in R.CONTROLS]          # the keys: control, assertion
    rows += [("ReconciliationExpected", name, F.RECONCILIATION_FORMULAS[name][0]) for name, _ in R.CONTROLS]
    rows += [("ReconciliationFound", name, F.RECONCILIATION_FORMULAS[name][1]) for name, _ in R.CONTROLS]
    rows += [("Census", name, text) for name, text in F.CENSUS_COLUMNS]
    rows += [("Unclassified", name, text) for name, text in F.UNCLASSIFIED_COLUMNS]
    rows += [("DoneStatuses", name, text) for name, text in F.DONE_STATUS_COLUMNS]
    rows += [("ParametersInForce", name, text) for name, text in F.PARAMETERS_IN_FORCE]
    rows += [("MetricItemsTitle", title, "") for title in F.METRICITEMS_TITLES]
    for block, columns in zip(("MetricItemsDone", "MetricItemsLeftOut", "MetricItemsCutoff"), F.METRICITEMS_BLOCKS):
        rows += [(block, name, text) for name, text in columns]
    rows += [("Controls", CONTROL_NAMES[0], F.AS_OF_CELL), ("Controls", CONTROL_NAMES[1], F.MODE_CELL),
             ("Controls", CONTROL_NAMES[2], F.TRANSITIONS_LAST_CELL), ("Controls", CONTROL_NAMES[3], F.PIS_LAST_CELL),
             ("Controls", CONTROL_NAMES[4], F.PREDICTABILITY_MODE_CELL)]
    rows += [("Dashboard", "ProgramTeam", ART), ("Dashboard", "AllTeams", D.ALL_TEAMS),
             ("Dashboard", "SelectorCell", D.SELECTOR_CELL), ("Dashboard", "NoPI", D.NO_PI)]
    return rows


def check_rows() -> list[tuple[str, ...]]:
    """The checks of the sheet DataChecks (Id, Title, Severity, three needs, Source): VBA writes them as the keys of the table."""
    assert all(len(c.needs) <= 3 for c in CHECKS), "DataChecks has three Need columns"
    return [(c.id, c.title, c.severity, *(c.needs + ("", "", ""))[:3], c.source) for c in CHECKS]


def chart_rows() -> list[tuple]:
    """One row per chart. The charts of the sheet IterationCharts come before the daily curve: its block of ChartData is the
    longest (one row per day), so it stays last."""
    rows = []
    for index, spec in enumerate(D.CHARTS[:-1]):
        rows.append(_chart_row(index, spec, D.CHARTS_SHEET))
    for index, spec in enumerate(D.ITERATION_CHARTS):
        rows.append(_chart_row(index, spec, D.ITERATION_CHARTS_SHEET))
    rows.append(_chart_row(len(D.CHARTS) - 1, D.CHARTS[-1], D.CHARTS_SHEET))
    return rows


def _chart_row(index: int, spec: D.ChartSpec, sheet: str) -> tuple:
    if sheet == D.ITERATION_CHARTS_SHEET:
        point = D.iteration_point(spec.column, "{iteration}", "{name}")      # one series per team of the Features
        series = ""
    elif spec.axis == D.DAY_AXIS:
        point = D.day_point("{iteration}", "{name}")                # {iteration} is the day here
        series = ""
    elif spec.series == D.STATUS_SERIES:
        point = D.status_point("{iteration}", "{name}")              # {iteration} is the PI here; one series per status
        series = D.STATUSES_TOKEN                                    # the statuses come from the sheet Workflows when the charts are built
    elif spec.axis == D.PI_AXIS:
        point = D.pi_point("{iteration}")                           # {iteration} is the PI here; one series, ART
        series = ",".join(D.series_names(spec))
    elif spec.series == D.TEAMS:
        point = D.team_point(spec.column, "{iteration}", "{name}")
        series = ""
    else:
        point = D.bucket_point("{series}", "{iteration}")
        series = ",".join(D.series_names(spec))
    row, column = D.chart_anchor(index)
    return (spec.name, spec.title, spec.chart_type, spec.axis_format or "", series, point,
            D.note_formula("{range}", spec.reason, spec.zeros), row, column, D.ROWS_PER_CHART, D.COLUMNS_PER_CHART,
            spec.axis, D.band_text(spec), sheet, "yes" if spec.labels else "")        # text, not a Boolean ("Vrai" on a French PC)


def team_names(project_rows: list[tuple[str, str, str]]) -> list[str]:
    """Test mirror of modFlowBuild.TeamNames (VBA computes the teams itself): ART (the Program), then the teams of the Team projects."""
    teams = [ART]
    for key, role, team in project_rows:
        team = (team or "").strip() or key.strip()          # an empty Team is the project key
        if role.strip().lower() == "team" and team and team not in teams:
            teams.append(team)
    return teams
