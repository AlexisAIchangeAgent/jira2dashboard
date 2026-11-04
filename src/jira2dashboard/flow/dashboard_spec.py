"""The dashboard: which charts, and the formulas of the data behind them (pure, nothing here touches Excel).

The sheet ChartData holds one block per chart: a row of series names, one row per period, one formula per point.
A point that is "not computable" is the Excel error #N/A, so the chart leaves a gap there and never shows a 0.
sheets.py builds the sheets and the charts from this description.
"""
from dataclasses import dataclass

from jira2dashboard.flow.formulas import BUCKETS

ALL_TEAMS = "All teams"
SELECTOR = "dashTeam"                         # name of the cell that picks the team of the Distribution chart
SELECTOR_CELL = "B2"

XL_COLUMN_CLUSTERED, XL_LINE_MARKERS, XL_COLUMN_STACKED_100 = 51, 65, 53

TEAMS, BUCKET_SERIES = "team", "bucket"


@dataclass(frozen=True)
class ChartSpec:
    name: str                       # name of the chart in the sheet, and of its PNG file
    title: str
    column: str | None              # the Metrics column drawn (the Distribution chart draws DistDone<bucket>)
    chart_type: int
    axis_format: str | None = None
    series: str = TEAMS             # one series per team, or one per work type bucket
    reason: str = ""                # shown under the chart when points are not computable


CHARTS = (
    ChartSpec("velocity", "Velocity - items done per period", "Velocity", XL_COLUMN_CLUSTERED),
    ChartSpec("flow_time", "Flow Time - median, in days", "FlowTimeMedian", XL_LINE_MARKERS),
    ChartSpec("flow_load", "Flow Load - average items in process", "FlowLoadAverage", XL_LINE_MARKERS),
    ChartSpec("flow_efficiency", "Flow Efficiency - active time / (active + wait)", "FlowEfficiency", XL_LINE_MARKERS,
              "0%"),
    ChartSpec("flow_distribution", "Flow Distribution - items done by work type (team of cell B2)", None,
              XL_COLUMN_STACKED_100, "0%", series=BUCKET_SERIES),
    ChartSpec("flow_predictability", "Flow Predictability", "FlowPredictability", XL_LINE_MARKERS, "0%",
              reason="it needs the PI objectives, which are not in Jira"),
)

COLUMNS_PER_CHART, ROWS_PER_CHART = 8, 18       # size of the area of a chart on the Dashboard sheet
CHART_SPACING = 20                              # rows between two charts (chart, then its note)
CHARTS_PER_ROW = 2


def block_top(index: int, periods: int) -> int:
    """First row of the block of chart `index` in ChartData: title, series names, then one row per period."""
    return 1 + index * (periods + 4)


def chart_anchor(index: int) -> tuple[int, int]:
    """(row, column) of the top-left cell of chart `index` on the Dashboard sheet."""
    return 4 + (index // CHARTS_PER_ROW) * CHART_SPACING, 1 + (index % CHARTS_PER_ROW) * (COLUMNS_PER_CHART + 1)


def team_point(column: str, period_cell: str, team_cell: str) -> str:
    """One point of a team series: the figure of Metrics for this period and team, or #N/A."""
    return (f'=LET(v,XLOOKUP(1,(Metrics[Period]={period_cell})*(Metrics[Team]={team_cell}),Metrics[{column}],""),'
            "IF(ISNUMBER(v),v,NA()))")


def bucket_point(bucket: str, period_cell: str) -> str:
    """One point of the Distribution chart: the share of one work type among the items done in the period, for the
    team picked in the selector, or for all teams together (each team's share weighted by its velocity)."""
    column = f"Metrics[DistDone{bucket}]"
    one_team = f'XLOOKUP(1,(Metrics[Period]={period_cell})*(Metrics[Team]={SELECTOR}),{column},"")'
    all_teams = (f"SUMPRODUCT((Metrics[Period]={period_cell})*Metrics[Velocity]*IFERROR(1*{column},0))"
                 f"/SUMIFS(Metrics[Velocity],Metrics[Period],{period_cell})")
    return f'=IFERROR(IF({SELECTOR}="{ALL_TEAMS}",{all_teams},1*{one_team}),NA())'


def note_formula(data_range: str, reason: str) -> str:
    """The text under a chart: how many points are not computable (gaps in the chart)."""
    because = f'&" - {reason}"' if reason else ""
    return (f'=SUMPRODUCT(--ISNA({data_range}))&" of "&ROWS({data_range})*COLUMNS({data_range})'
            f'&" points are not computable (gaps)"{because}')


def series_names(spec: ChartSpec) -> tuple[str, ...] | None:
    """The fixed series names (None when they are the teams, which come from the sheet Projects)."""
    return BUCKETS if spec.series == BUCKET_SERIES else None
