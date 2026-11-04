"""The dashboard: which charts, and the formulas of the data behind them (pure, nothing here touches Excel).

The sheet ChartData holds one block per chart: a row of series names, one row per period, one formula per point.
A point that is "not computable" is the Excel error #N/A, so the chart leaves a gap there and never shows a 0.
blueprint.py writes this description into the sheet ChartSpec; modFlowBuild.bas builds ChartData and the charts from it.
"""
from dataclasses import dataclass

from jira2dashboard.flow.formulas import BUCKETS

ALL_TEAMS = "All teams"
SELECTOR = "dashTeam"                         # name of the cell that picks the team of the Distribution chart
SELECTOR_CELL = "B2"
STATUS = "dashStatus"                          # name of the cell where the extraction tells what it does
STATUS_CELL = "A3"
BAR_CELL, CANCEL_CELL = "J3", "M1"            # the progress bar (rectangles pbBack, pbFill) and the button Cancel

XL_COLUMN_CLUSTERED, XL_LINE_MARKERS, XL_COLUMN_STACKED_100, XL_LINE = 51, 65, 53, 4

TEAMS, BUCKET_SERIES = "team", "bucket"
PERIOD_AXIS, DAY_AXIS = "period", "day"          # one point per period (table Metrics) or per day (table DailyLoad)
BAND_NAMES = ("Target band low", "Target band high")


@dataclass(frozen=True)
class ChartSpec:
    name: str                       # name of the chart in the sheet, and of its PNG file
    title: str
    column: str | None              # the Metrics column drawn (the Distribution chart draws DistDone<bucket>)
    chart_type: int
    axis_format: str | None = None
    series: str = TEAMS             # one series per team, or one per work type bucket
    reason: str = ""                # shown under the chart when points are not computable
    axis: str = PERIOD_AXIS         # what a row of the block of ChartData is: a period or a day
    band: tuple[float, float] | None = None     # a target band drawn as two constant series (proposal: SAFe 80-100 %)


CHARTS = (
    ChartSpec("velocity", "Velocity - items done per period", "Velocity", XL_COLUMN_CLUSTERED),
    ChartSpec("flow_time", "Flow Time - median, in days", "FlowTimeMedian", XL_LINE_MARKERS),
    ChartSpec("flow_time_p85", "Flow Time - 85th percentile, in days", "FlowTimeP85", XL_LINE_MARKERS),
    ChartSpec("flow_load", "Flow Load - average items in process", "FlowLoadAverage", XL_LINE_MARKERS),
    ChartSpec("flow_efficiency", "Flow Efficiency - active time / (active + wait)", "FlowEfficiency",
              XL_COLUMN_CLUSTERED, "0%"),
    ChartSpec("flow_distribution", "Flow Distribution - items done by work type (team of the selector)", None,
              XL_COLUMN_STACKED_100, "0%", series=BUCKET_SERIES),
    ChartSpec("flow_predictability", "Flow Predictability", "FlowPredictability", XL_LINE_MARKERS, "0%",
              reason="it needs the PI objectives, which are not in Jira", band=(0.8, 1.0)),
    ChartSpec("flow_load_daily", "Flow Load - items in process at the end of each day", None, XL_LINE, axis=DAY_AXIS,
              reason="the day is after the as-of date"),          # last: its block has one row per day, not per period
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


def day_point(day_cell: str, team_cell: str) -> str:
    """One point of the daily Flow Load curve: the snapshot of DailyLoad for this day and team, or #N/A (no snapshot yet)."""
    return (f'=LET(v,XLOOKUP(1,(DailyLoad[Date]={day_cell})*(DailyLoad[Team]={team_cell}),DailyLoad[Load],""),'
            "IF(ISNUMBER(v),v,NA()))")


def band_text(spec: ChartSpec) -> str:
    """The band as the text that modFlowBuild reads ("0.8;1", a dot whatever the language of Excel), or empty."""
    return f"{spec.band[0]};{spec.band[1]}" if spec.band else ""


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
