"""The dashboard: which charts, and the formulas of the data behind them (pure, nothing here touches Excel).

The sheet ChartData holds one block per chart: a row of series names, one row per iteration, one formula per point.
A point that is "not computable" is the Excel error #N/A, so the chart leaves a gap there and never shows a 0.
blueprint.py writes this description into the sheet ChartSpec; modFlowBuild.bas builds ChartData and the charts from it.
"""
from dataclasses import dataclass

from jira2dashboard import sheet_names as S
from jira2dashboard.flow.formulas import BUCKETS
from jira2dashboard.flow.model import ART
from jira2dashboard.flow.refs import HALF

ALL_TEAMS = "All teams"
CHARTS_SHEET = "Charts"                       # the sheet of the charts and of the team selector, right of the Dashboard
SELECTOR = "dashTeam"                         # name of the cell that picks the team of the Distribution chart
SELECTOR_CELL = "B2"                          # on the sheet Charts
STATUS = "dashStatus"                          # name of the cell where the extraction tells what it does (its headline)
# Top left, one under the other (AL's order): the progress bar, Prompt for Rovo, Load settings CSV, Read the Jira
# configuration, Propose values from Jira, then the three steps (AL, 2026-11-22: correct and compute again without reading Jira):
# 1. Extract from Jira, 2. Compute the metrics, 3. Draw the charts, Cancel, and last Debug prompt for Rovo (AL, 2026-11-26); then the status,
# so that no text of the status runs under a button.
BAR_CELL = "A2"                               # the progress bar (rectangles pbBack, pbFill)
ROVO_CELL = "A3"                              # the button that writes the prompt for Rovo
ROVO_LOAD_CELL = "A5"                         # the button that loads the settings CSV (the answer of Rovo)
CONFIG_CELL = "A7"                            # the button that reads the Jira configuration into proposal sheets
PROPOSE_CELL = "A9"                           # the button that proposes values in the parameter sheets (empty cells only)
EXTRACT_CELL = "A11"                          # the button 1. Extract from Jira
COMPUTE_CELL = "A13"                          # the button 2. Compute the metrics
DRAW_CELL = "A15"                             # the button 3. Draw the charts
CANCEL_CELL = "A17"                           # the button Cancel
ROVO_DEBUG_CELL = "A19"                       # the button that writes the debug prompt for Rovo (the files of rovo-debug), last (AL, 2026-11-26)
STATUS_CELL = "A22"
STATUS_ROWS = 20                              # rows of fields under it (A23:A42): name in A, value in B (modActions.ShowProgress)
BUTTON_WIDTH, BUTTON_HEIGHT, BAR_HEIGHT = 190, 26, 14

XL_COLUMN_CLUSTERED, XL_LINE_MARKERS, XL_COLUMN_STACKED_100, XL_LINE = 51, 65, 53, 4

TEAMS, BUCKET_SERIES, ART_SERIES = "team", "bucket", "art"     # one series per team, per work type, or ART alone
STATUS_SERIES = "status"                       # one series per Active / Wait status of the Feature workflow (read from the sheet Workflows)
NO_STATUS = "(no status)"                      # the one status of PIStatusTime when the Feature workflow has no Active or Wait status (modFlowBuild.NO_STATUS)
STATUSES_TOKEN = "{statuses}"                  # the plan column SeriesNames of such a chart: modFlowBuild.BuildChart replaces it by StatusNames
ITERATION_AXIS, DAY_AXIS, PI_AXIS = "iteration", "day", "pi"   # one point per iteration (table Metrics), per day (table
#                                                                 DailyLoad) or per PI (rows of the sheet PIs)
NO_PI = "(no PI in the sheet PIs)"                             # the one row of a PI chart while the sheet PIs is empty
BAND_NAMES = ("Target band low", "Target band high")
BAND_MARKER = -4115         # xlMarkerStyleDash: a band line has a long bar on each point, so one PI still shows it (AL)


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
    labels: bool = False            # the value above each bar of the first series (the ART), so that a 0 is seen: no bar has no height
    zeros: bool = False             # the note counts the points that are 0 (no bar) instead of the gaps
    causes: tuple[str, str] | None = None   # (table, column) of the cause of each point: the note names the most frequent one (AL, 2026-11-24)


EXTRACT_CAPTION = "1. Extract from Jira"              # the captions of the Dashboard buttons (flow/sheets.py; modActions has the three steps)
COMPUTE_CAPTION = "2. Compute the metrics"
DRAW_CAPTION = "3. Draw the charts"
CANCEL_CAPTION = "Cancel"
CONFIG_CAPTION = "Read the Jira configuration"
PROPOSE_CAPTION = "Propose values from Jira"
ROVO_CAPTION = "Prompt for Rovo"
ROVO_LOAD_CAPTION = "Load settings CSV"
ROVO_DEBUG_CAPTION = "Debug prompt for Rovo"
STEPS_CELL = "T1"           # the first steps (manual.first_steps), right of the status rows

# AL, 2026-11-17: an iteration where nothing was done (every Feature deferred, say) has Velocity 0 (no bar) and no Flow Time or Flow
# Efficiency (a gap: never 0, never a guess). The value is written above the bar and the note under the chart says why.
NO_ITEM_DONE = "no item was done in these iterations, so there is nothing to measure"
# AL, 2026-11-22: the note names every cause of a gap, the start status first (a setting to fix; WhyNotInFlowTime of the sheet MetricItems says it per ticket)
FLOW_TIME_REASON = ("no item was done in these iterations, or every done item was left out: no start status (IsStart, the sheet Workflows_from_user), "
                    "never in the start status, back to the backlog after it, or reopened (WhyNotInFlowTime, the sheet MetricItems_from_excel)")
EFFICIENCY_REASON = ("no item was done in these iterations, or every done item was left out (no start status: see Flow Time), or no status is "
                     "declared Wait (the sheet Workflows_from_user): Jira cannot say which in-progress statuses wait")
LOAD_REASON = "the workflow of the team has no start status (IsStart, the sheet Workflows_from_user), or the iteration ends after the as-of date"
DAILY_LOAD_REASON = "the day is after the as-of date, or the workflow of the team has no start status (IsStart, the sheet Workflows_from_user)"
PI_STATUS_REASON = "no Feature was released in the PI, or the PI is not in the sheet PIs_from_user"
ZERO_NOTE = "no item was done (no bar)"

CHARTS = (
    ChartSpec("velocity", "Velocity - items done per iteration", "Velocity", XL_COLUMN_CLUSTERED, labels=True, zeros=True),
    ChartSpec("flow_time", "Flow Time - median, in days", "FlowTimeMedian", XL_LINE_MARKERS, reason=FLOW_TIME_REASON,
              causes=("FlowTimeCheck", "Cause")),
    ChartSpec("flow_time_p85", "Flow Time - 85th percentile, in days", "FlowTimeP85", XL_LINE_MARKERS, reason=FLOW_TIME_REASON,
              causes=("FlowTimeCheck", "Cause")),
    ChartSpec("flow_load", "Flow Load - average items in process", "FlowLoadAverage", XL_LINE_MARKERS, reason=LOAD_REASON,
              causes=("FlowLoadCheck", "Reason")),
    ChartSpec("flow_efficiency", "Flow Efficiency - active time / (active + wait)", "FlowEfficiency",
              XL_COLUMN_CLUSTERED, "0%", reason=EFFICIENCY_REASON, causes=("FlowEfficiencyCheck", "Cause")),
    ChartSpec("flow_distribution", "Flow Distribution - items done by work type (team of the selector)", None,
              XL_COLUMN_STACKED_100, "0%", series=BUCKET_SERIES),
    ChartSpec("flow_predictability", "Flow Predictability", "FlowPredictability", XL_LINE_MARKERS, "0%",
              reason="it needs the PI objectives, which are not in Jira", band=(0.8, 1.0)),
    # AL's variant (2026-11-12): the Features done before the end of the PI over the Features planned in it (flow/metrics.py)
    ChartSpec("feature_predictability", "Feature Predictability - features released in the PI / features engaged at its start",
              None, XL_COLUMN_CLUSTERED, "0%", series=ART_SERIES, axis=PI_AXIS, band=(0.8, 1.0),
              reason="no feature planned in the PI, or the PI is not in the sheet PIs_from_user"),
    # AL, 2026-11-17: where the cycle time goes. One bar per PI, one colour per status; the bars add up to 100 %. The purpose is to
    # see the statuses that take most of it (flow/pi_status_time.py, table PIStatusTime).
    ChartSpec("pi_status_time", "Cycle time by status - share of the average cycle time of the Features released in each PI", None,
              XL_COLUMN_STACKED_100, "0%", series=STATUS_SERIES, axis=PI_AXIS,
              reason=PI_STATUS_REASON),
    ChartSpec("flow_load_daily", "Flow Load - items in process at the end of each day", None, XL_LINE, axis=DAY_AXIS,
              reason=DAILY_LOAD_REASON),          # last: its block has one row per day, not per iteration
)

# The sheet IterationCharts (AL, 2026-11-16): the Features of the ART iteration by iteration (table IterationFeatures), one series per
# team of the Features plus the ART as a whole; the PI figures of Feature Predictability are summed from these counts.
ITERATION_CHARTS_SHEET = "IterationCharts"
ITERATION_CHARTS = (
    ChartSpec("features_engaged", "Features engaged at the start of each iteration", "Engaged", XL_COLUMN_CLUSTERED),
    ChartSpec("features_added", "Features added to the PI during each iteration", "Added", XL_COLUMN_CLUSTERED),
    ChartSpec("features_deferred", "Features deferred out of the PI during each iteration", "Deferred", XL_COLUMN_CLUSTERED),
    ChartSpec("features_released", "Features released in each iteration", "Released", XL_COLUMN_CLUSTERED),
    ChartSpec("features_carried_over", "Features carried over to the next iteration", "CarryOver", XL_COLUMN_CLUSTERED),
)

COLUMNS_PER_CHART, ROWS_PER_CHART = 8, 18       # size of the area of a chart on the sheet Charts
CHART_SPACING = 20                              # rows between two charts (chart, then its note)
CHARTS_PER_ROW = 2


def block_top(index: int, rows: int) -> int:
    """First row of the block of chart `index` in ChartData: title, series names, then one row per iteration or PI. `rows` is
    the larger of the two counts (at least 1 PI row), as in modFlowBuild.BuildCharts; the daily curve, last, is longer."""
    return 1 + index * (rows + 4)


def chart_anchor(index: int) -> tuple[int, int]:
    """(row, column) of the top-left cell of chart `index` on the sheet Charts (rows 1-2: its title and the selector)."""
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


def iteration_point(column: str, iteration_cell: str, team_cell: str) -> str:
    """One point of a chart of IterationCharts: the count of IterationFeatures for this iteration and team (the rows of the PIs that
    hold the iteration are added: a Feature carries one PI at a time)."""
    return (f"=SUMIFS(IterationFeatures[{column}],IterationFeatures[[Iteration-label]],{iteration_cell},"
            f"IterationFeatures[Team],{team_cell})")


def pi_point(pi_cell: str) -> str:
    """One point of the Feature Predictability chart (ART): the FeaturePredictability of the PI of the row in the table
    PIMetrics (formulas.PIMETRICS_COLUMNS), or #N/A when it is not computable (no Feature planned in the PI)."""
    return (f'=LET(v,XLOOKUP({pi_cell},PIMetrics[PI],PIMetrics[FeaturePredictability],""),'
            "IF(ISNUMBER(v),v,NA()))")


def status_point(pi_cell: str, name_cell: str) -> str:
    """One point of the Cycle time by status chart: the Share of the status (the series name, in the header cell) in the PI of the row,
    from the table PIStatusTime (formulas.PISTATUSTIME_COLUMNS), or #N/A when it is not computable (no Feature released in the PI)."""
    return (f"=LET(v,XLOOKUP(1,(PIStatusTime[PI]={pi_cell})*(PIStatusTime[Status]={name_cell}),PIStatusTime[Share],\"\"),"
            "IF(ISNUMBER(v),v,NA()))")


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


def most_often(causes: list[str]) -> tuple[str, int, int] | None:
    """(the most frequent cause, its count, the rows with a cause) of the rows of a check table, the first met when two are as frequent (SORTBY
    keeps the order of UNIQUE); None when no row has one."""
    named = [c for c in causes if c]
    if not named:
        return None
    counts = {c: named.count(c) for c in dict.fromkeys(named)}
    top = max(counts, key=counts.get)
    return top, counts[top], len(named)


def most_often_text(cause: str, count: int, rows: int, table: str) -> str:
    return f" - most often ({count} of {rows} rows of {S.of(table)}): {cause}"


def note_formula(data_range: str, reason: str, zeros: bool = False, causes: tuple[str, str] | None = None) -> str:
    """The text under a chart: how many points are not computable (gaps in the chart), or with `zeros` how many are 0 (no bar). With `causes`
    (a check table and its column of causes) the most frequent cause of its rows instead of the list of every possible one, which stays
    when no row has a cause; nothing when there is no gap. LET names: no `r` or `c` (Excel reads them as R1C1)."""
    because = f'&" - {reason}"' if reason else ""
    if causes:
        col = f"{causes[0]}[{causes[1]}]"
        top = (f'LET(rs,FILTER({col},{col}<>""),un,UNIQUE(rs),ct,MMULT(--(un=TRANSPOSE(rs)),SEQUENCE(ROWS(rs),1,1,0)),'
               f'" - most often ("&MAX(ct)&" of "&ROWS(rs)&" rows of {S.of(causes[0])}): "&INDEX(SORTBY(un,ct,-1),1))')
        return (f'=LET(n,SUMPRODUCT(--ISNA({data_range})),n&" of "&ROWS({data_range})*COLUMNS({data_range})&" points are not computable (gaps)"'
                f'&IF(n=0,"",IFERROR({top},"{" - " + reason if reason else ""}")))')
    if zeros:
        return (f'=COUNTIF({data_range},0)&" of "&ROWS({data_range})*COLUMNS({data_range})'
                f'&" points are 0 - {ZERO_NOTE}"{because}')
    return (f'=SUMPRODUCT(--ISNA({data_range}))&" of "&ROWS({data_range})*COLUMNS({data_range})'
            f'&" points are not computable (gaps)"{because}')


def series_names(spec: ChartSpec) -> tuple[str, ...] | None:
    """The fixed series names (None when they are the teams, which come from the sheet Projects)."""
    if spec.series == ART_SERIES:
        return (ART,)
    return BUCKETS if spec.series == BUCKET_SERIES else None
