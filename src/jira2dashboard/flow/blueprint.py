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
from jira2dashboard.flow.checks import CHECKS
from jira2dashboard.flow.model import ART

SHEET_FORMULAS = "FlowFormulas"
SHEET_CHARTS = "ChartSpec"
SHEET_CHECKS = "ChecksSpec"
FORMULAS_HEADER = ("Table", "Column", "Formula")
CHECKS_HEADER = C.DATACHECKS_KEYS
CHARTS_HEADER = ("Name", "Title", "ChartType", "AxisFormat", "SeriesNames", "PointFormula", "NoteFormula",
                 "AnchorRow", "AnchorColumn", "Rows", "Columns", "Axis", "Band")

CONTROL_NAMES = ("AsOfUtc", "cfg_FlowLoadMode")        # the names of the control cells (sheet Controls, column B)


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
    rows += [("Metrics", name, text) for name, text in F.METRICS_COLUMNS]
    rows += [("Controls", CONTROL_NAMES[0], F.AS_OF_CELL), ("Controls", CONTROL_NAMES[1], F.MODE_CELL)]
    rows += [("Dashboard", "ProgramTeam", ART), ("Dashboard", "AllTeams", D.ALL_TEAMS),
             ("Dashboard", "SelectorCell", D.SELECTOR_CELL)]
    return rows


def check_rows() -> list[tuple[str, ...]]:
    """The checks of the sheet DataChecks (Id, Title, Severity, three needs, Source): VBA writes them as the keys of the table."""
    assert all(len(c.needs) <= 3 for c in CHECKS), "DataChecks has three Need columns"
    return [(c.id, c.title, c.severity, *(c.needs + ("", "", ""))[:3], c.source) for c in CHECKS]


def chart_rows() -> list[tuple]:
    rows = []
    for index, spec in enumerate(D.CHARTS):
        if spec.axis == D.DAY_AXIS:
            point = D.day_point("{period}", "{name}")                # {period} is the day here
            series = ""
        elif spec.series == D.TEAMS:
            point = D.team_point(spec.column, "{period}", "{name}")
            series = ""
        else:
            point = D.bucket_point("{series}", "{period}")
            series = ",".join(D.series_names(spec))
        row, column = D.chart_anchor(index)
        rows.append((spec.name, spec.title, spec.chart_type, spec.axis_format or "", series, point,
                     D.note_formula("{range}", spec.reason), row, column, D.ROWS_PER_CHART, D.COLUMNS_PER_CHART,
                     spec.axis, D.band_text(spec)))
    return rows


def team_names(project_rows: list[tuple[str, str, str]]) -> list[str]:
    """Test mirror of modFlowBuild.TeamNames (VBA computes the teams itself): ART (the Program), then the teams of the Team projects."""
    teams = [ART]
    for _key, role, team in project_rows:
        if role.strip().lower() == "team" and team.strip() and team.strip() not in teams:
            teams.append(team.strip())
    return teams
