"""The dashboard: which charts, and the formulas of the data behind them (pure, nothing here touches Excel).

The sheet ChartData holds one block per chart: a row of series names, one row per iteration, one formula per point.
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
CONFIG_CELL = "O1"                            # the button that reads the Jira configuration into proposal sheets
ROVO_CELL = "R1"                              # the button that writes the prompt for Rovo

XL_COLUMN_CLUSTERED, XL_LINE_MARKERS, XL_COLUMN_STACKED_100, XL_LINE = 51, 65, 53, 4

TEAMS, BUCKET_SERIES = "team", "bucket"
ITERATION_AXIS, DAY_AXIS = "iteration", "day"          # one point per iteration (table Metrics) or per day (table DailyLoad)
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
    axis: str = ITERATION_AXIS         # what a row of the block of ChartData is: an iteration or a day
    band: tuple[float, float] | None = None     # a target band drawn as two constant series (proposal: SAFe 80-100 %)


EXTRACT_CAPTION = "Extract from Jira and compute"     # the captions of the Dashboard buttons (flow/sheets.py)
CANCEL_CAPTION = "Cancel"
CONFIG_CAPTION = "Read the Jira configuration"
ROVO_CAPTION = "Prompt for Rovo"
STEPS_CELL = "T1"           # the first steps (manual.first_steps): beside the charts; a rebuild clears only A4:R400

CHARTS = (
    ChartSpec("velocity", "Velocity - items done per iteration", "Velocity", XL_COLUMN_CLUSTERED),
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
              reason="the day is after the as-of date"),          # last: its block has one row per day, not per iteration
)

COLUMNS_PER_CHART, ROWS_PER_CHART = 8, 18       # size of the area of a chart on the Dashboard sheet
CHART_SPACING = 20                              # rows between two charts (chart, then its note)
CHARTS_PER_ROW = 2


def block_top(index: int, iterations: int) -> int:
    """First row of the block of chart `index` in ChartData: title, series names, then one row per iteration."""
    return 1 + index * (iterations + 4)


def chart_anchor(index: int) -> tuple[int, int]:
    """(row, column) of the top-left cell of chart `index` on the Dashboard sheet."""
    return 4 + (index // CHARTS_PER_ROW) * CHART_SPACING, 1 + (index % CHARTS_PER_ROW) * (COLUMNS_PER_CHART + 1)


def team_point(column: str, iteration_cell: str, team_cell: str) -> str:
    """One point of a team series: the figure of Metrics for this iteration and team, or #N/A."""
    return (f'=LET(v,XLOOKUP(1,(Metrics[[Iteration-label]]={iteration_cell})*(Metrics[Team]={team_cell}),Metrics[{column}],""),'
            "IF(ISNUMBER(v),v,NA()))")


def day_point(day_cell: str, team_cell: str) -> str:
    """One point of the daily Flow Load curve: the snapshot of DailyLoad for this day and team, or #N/A (no snapshot yet).
    COUNTIFS and SUMIFS, not an array lookup: one point per day and team would scan the whole table each time."""
    keys = f"DailyLoad[Date],{day_cell},DailyLoad[Team],{team_cell}"
    return f'=IF(COUNTIFS({keys},DailyLoad[Load],">=0")=0,NA(),SUMIFS(DailyLoad[Load],{keys}))'


def band_text(spec: ChartSpec) -> str:
    """The band as the text that modFlowBuild reads ("0.8;1", a dot whatever the language of Excel), or empty."""
    return f"{spec.band[0]};{spec.band[1]}" if spec.band else ""


def bucket_point(bucket: str, iteration_cell: str) -> str:
    """One point of the Distribution chart: the share of one work type among the items done in the iteration, for the
    team picked in the selector, or for all teams together (each team's share weighted by its velocity)."""
    column = f"Metrics[DistDone{bucket}]"
    one_team = f'XLOOKUP(1,(Metrics[[Iteration-label]]={iteration_cell})*(Metrics[Team]={SELECTOR}),{column},"")'
    all_teams = (f"SUMPRODUCT((Metrics[[Iteration-label]]={iteration_cell})*Metrics[Velocity]*IFERROR(1*{column},0))"
                 f"/SUMIFS(Metrics[Velocity],Metrics[[Iteration-label]],{iteration_cell})")
    return f'=IFERROR(IF({SELECTOR}="{ALL_TEAMS}",{all_teams},1*{one_team}),NA())'


def note_formula(data_range: str, reason: str) -> str:
    """The text under a chart: how many points are not computable (gaps in the chart)."""
    because = f'&" - {reason}"' if reason else ""
    return (f'=SUMPRODUCT(--ISNA({data_range}))&" of "&ROWS({data_range})*COLUMNS({data_range})'
            f'&" points are not computable (gaps)"{because}')


def series_names(spec: ChartSpec) -> tuple[str, ...] | None:
    """The fixed series names (None when they are the teams, which come from the sheet Projects)."""
    return BUCKETS if spec.series == BUCKET_SERIES else None
