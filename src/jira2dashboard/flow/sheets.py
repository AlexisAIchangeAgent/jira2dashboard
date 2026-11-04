"""Write the flow sheets into a workbook (needs Excel): tables, formula columns, control cells.

The raw sheets Issues and StatusIntervals are written by VBA (modExtract). This module turns them into tables, adds
the helper columns and builds ItemFlow, DailyLoad and Metrics as formulas (formulas.py). Nothing is computed here.
"""
import datetime as dt

from jira2dashboard.excel import com as R
from jira2dashboard.flow import dashboard_spec as D
from jira2dashboard.flow import formulas as F

EXCEL_ERRORS = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
                -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!", -2146826254: "#CALC!"}


def _used_rows(ws) -> int:
    return ws.Cells(ws.Rows.Count, 1).End(-4162).Row                      # -4162 = xlUp


def _table(ws, name: str, last_row: int, columns: int):
    """A table over A1:<last column><last_row>, header in row 1 (positional arguments: xlSrcRange, range, -, xlYes)."""
    rng = ws.Range(ws.Cells(1, 1), ws.Cells(max(last_row, 2), columns))
    table = ws.ListObjects.Add(1, rng, None, 1)
    table.Name = name
    return table


class FormulaRejected(ValueError):
    """Excel refused a formula: the message names the table, the column and the text."""


def _fill(table, name: str, formula: str) -> None:
    try:
        table.ListColumns(name).DataBodyRange.Formula2 = formula
    except R.pywintypes.com_error as err:
        columns = [table.ListColumns(i).Name for i in range(1, table.ListColumns.Count + 1)]
        raise FormulaRejected(f"{table.Name}.{name} refused by Excel ({err.excepinfo[5] if err.excepinfo else err}): "
                              f"{formula}\ncolumns: {columns}") from None


def read_table(wb, sheet: str, name: str) -> tuple[list[str], list[list]]:
    """Header names and the raw values (Value2: dates are serial numbers, errors are negative integers)."""
    table = wb.Worksheets(sheet).ListObjects(name)
    header = [h for h in table.HeaderRowRange.Value2[0]]
    body = table.DataBodyRange.Value2
    return header, [list(row) for row in body]


def error_cells(wb) -> dict[str, int]:
    """How many cells of each flow table show an Excel error, by column (a typo in a formula shows here)."""
    found: dict[str, int] = {}
    for sheet, name in (("StatusIntervals", "StatusIntervals"), ("ItemFlow", "ItemFlow"), ("DailyLoad", "DailyLoad"),
                        ("Metrics", "Metrics")):
        header, rows = read_table(wb, sheet, name)
        for row in rows:
            for column, value in zip(header, row):
                if isinstance(value, int) and value in EXCEL_ERRORS:
                    key = f"{name}.{column} {EXCEL_ERRORS[value]}"
                    found[key] = found.get(key, 0) + 1
    return found


def build_flow_sheets(wb) -> None:
    """Create everything after the VBA extraction. Needs the sheets Issues, StatusIntervals, Workflows, Projects,
    Periods (with its rows) and Config."""
    ws_issues, ws_si = wb.Worksheets("Issues"), wb.Worksheets("StatusIntervals")
    issue_count = _used_rows(ws_issues) - 1
    si_count = _used_rows(ws_si) - 1
    _table(ws_issues, "Issues", issue_count + 1, len(F.ISSUES))
    si = _table(ws_si, "StatusIntervals", si_count + 1, len(F.INTERVALS_RAW))

    for name, _formula in F.INTERVAL_HELPERS:
        si.ListColumns.Add().Name = name

    item = _create_table(wb, "ItemFlow", F.TABLES["ItemFlow"], issue_count)
    teams = _team_rows(wb)
    days = _days(wb)
    daily = _create_daily_load(wb, days, teams)
    metrics = _create_metrics(wb, teams)
    wb.Names.Add("AsOfUtc", "=Metrics!$B$1")                    # the control cells of the Metrics sheet
    wb.Names.Add("cfg_FlowLoadMode", "=Metrics!$B$2")
    metrics_sheet = wb.Worksheets("Metrics")
    metrics_sheet.Range("A1").Value = "AsOfUtc"
    metrics_sheet.Range("B1").Formula2 = F.AS_OF_CELL
    metrics_sheet.Range("A2").Value = "FlowLoadMode"
    metrics_sheet.Range("B2").Formula2 = F.MODE_CELL

    for name, formula in F.INTERVAL_HELPERS:
        _fill(si, name, formula)
    for name, formula in F.ITEMFLOW_COLUMNS:
        _fill(item, name, formula)
    _fill(daily, "Load", F.DAILYLOAD_LOAD)
    for name, formula in F.METRICS_COLUMNS:
        _fill(metrics, name, formula)
    wb.Application.CalculateFull()


def _create_table(wb, name: str, columns: tuple[str, ...], rows: int):
    ws = R.add_sheet_after_last(wb, name)
    R.write_block(ws, 1, 1, [list(columns)])
    return _table(ws, name, rows + 1, len(columns))


def _team_rows(wb) -> list[tuple[str, str]]:
    """(team name, formula giving it): ART first, then the Team projects of the sheet Projects, one row per team."""
    ws = wb.Worksheets("Projects")
    rows = ws.Range(ws.Cells(2, 1), ws.Cells(_used_rows(ws), 3)).Value2
    teams = [("ART", '="ART"')]
    seen = {"ART"}
    for offset, (_key, role, team) in enumerate(rows, start=2):
        if role == "Team" and team and team not in seen:
            seen.add(team)
            teams.append((team, f"=Projects!$C${offset}"))
    return teams


def _period_rows(wb) -> list[tuple[int, float, float]]:
    ws = wb.Worksheets("Periods")
    last = _used_rows(ws)
    return [(r, ws.Cells(r, 2).Value2, ws.Cells(r, 3).Value2) for r in range(2, last + 1)] if last >= 2 else []


def _days(wb) -> list[float]:
    """Serial numbers of the days from the first period start to the last period end (exclusive)."""
    periods = _period_rows(wb)
    first, last = min(int(p[1]) for p in periods), max(int(p[2]) for p in periods)
    return list(range(first, last))


def _create_daily_load(wb, days: list[float], teams: list[tuple[str, str]]):
    ws = R.add_sheet_after_last(wb, "DailyLoad")
    R.write_block(ws, 1, 1, [list(F.TABLES["DailyLoad"])])
    n = len(days) * len(teams)
    ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, 1)).Value2 = tuple((d,) for d in days for _ in teams)
    _write_formulas(ws.Range(ws.Cells(2, 2), ws.Cells(n + 1, 2)), [(t[1],) for _ in days for t in teams])
    return _table(ws, "DailyLoad", n + 1, 3)            # the table comes after the keys (see _write_formulas)


def _write_formulas(rng, rows: list[tuple]) -> None:
    """One formula per cell, in a plain range. Inside a table, a formula typed in one cell becomes the formula of
    the whole column (calculated column): the key columns must be written before the table is created. The cells
    are read back to prove they differ."""
    rng.Value = tuple(rows)
    back = rng.Formula
    flat = [cell for row in back for cell in row] if isinstance(back, tuple) else [back]
    wanted = [cell for row in rows for cell in row]
    if flat != wanted:
        raise FormulaRejected(f"formulas were not written cell by cell ({len(set(flat))} distinct instead of "
                              f"{len(set(wanted))})")


def _create_metrics(wb, teams: list[tuple[str, str]]):
    ws = R.add_sheet_after_last(wb, "Metrics")
    periods = _period_rows(wb)
    header_row = 4
    columns = F.TABLES["Metrics"]
    R.write_block(ws, header_row, 1, [list(columns)])
    n = len(periods) * len(teams)
    keys = []
    for row, _start, _end in periods:
        for _team, team_formula in teams:
            keys.append((f"=Periods!$A${row}", f"=Periods!$B${row}", f"=Periods!$C${row}", team_formula))
    _write_formulas(ws.Range(ws.Cells(header_row + 1, 1), ws.Cells(header_row + n, 4)), keys)
    rng = ws.Range(ws.Cells(header_row, 1), ws.Cells(header_row + n, len(columns)))
    table = ws.ListObjects.Add(1, rng, None, 1)
    table.Name = "Metrics"
    return table


def _address(ws, row: int, column: int, last_row: int | None = None, last_column: int | None = None) -> str:
    """'ChartData'!$B$5:$J$9 (absolute, with the sheet name), for series and formulas."""
    first = ws.Cells(row, column).Address
    if last_row is None:
        return f"{ws.Name}!{first}"
    return f"{ws.Name}!{first}:{ws.Cells(last_row, last_column).Address}"


def build_dashboard(wb) -> None:
    """The sheets ChartData (one block of formulas per chart) and Dashboard (the charts, the team selector).
    Needs the flow sheets (Metrics) and Periods. Nothing is computed here: every point is a formula."""
    teams, periods = _team_rows(wb), _period_rows(wb)
    data = R.add_sheet_after_last(wb, "ChartData")
    dash = R.add_sheet_after_last(wb, "Dashboard")
    _build_selector(wb, dash, data, teams)
    for index, spec in enumerate(D.CHARTS):
        top = D.block_top(index, len(periods))
        points = _write_block(data, spec, top, periods, teams)
        _add_chart(dash, spec, index, points)
    wb.Application.CalculateFull()


def _build_selector(wb, dash, data, teams) -> None:
    """Cell B2 of the Dashboard: a list (All teams + the teams of the sheet Projects) for the Distribution chart."""
    data.Range("AA1").Value = "Teams for the selector"
    _write_formulas(data.Range(data.Cells(2, 27), data.Cells(len(teams) + 2, 27)),
                    [(D.ALL_TEAMS,)] + [(t[1],) for t in teams])
    dash.Range("A1").Value = "Flow metrics"
    dash.Range("A1").Font.Bold = True
    dash.Range("A2").Value = "Team of the Distribution chart"
    dash.Range(D.SELECTOR_CELL).Value = D.ALL_TEAMS
    wb.Names.Add(D.SELECTOR, f"=Dashboard!${D.SELECTOR_CELL[0]}${D.SELECTOR_CELL[1:]}")
    # Validation.Add(Type = xlValidateList, AlertStyle = xlValidAlertStop, Operator = xlBetween, Formula1)
    dash.Range(D.SELECTOR_CELL).Validation.Add(3, 1, 1, f"=ChartData!$AA$2:$AA${len(teams) + 2}")
    dash.Range("E1").Value = "As of (UTC)"
    dash.Range("F1").Formula2 = "=AsOfUtc"
    R.put_en_us(dash.Range("F1"), "NumberFormat", "yyyy-mm-dd hh:mm")
    dash.Range("A:R").ColumnWidth = 11


def _write_block(data, spec, top: int, periods, teams):
    """Title, series names, period labels and one formula for every point. Returns the cells of the points."""
    series = D.series_names(spec)
    names = [(n,) for n in series] if series else [(t[1],) for t in teams]
    count, first = len(names), top + 2
    last = first + len(periods) - 1
    data.Cells(top, 1).Value = spec.title
    data.Cells(top, 1).Font.Bold = True
    data.Cells(top + 1, 1).Value = "Period"
    _write_formulas(data.Range(data.Cells(top + 1, 2), data.Cells(top + 1, count + 1)), [tuple(n[0] for n in names)])
    _write_formulas(data.Range(data.Cells(first, 1), data.Cells(last, 1)),
                    [(f"=Periods!$A${row}",) for row, _s, _e in periods])
    block = data.Range(data.Cells(first, 2), data.Cells(last, count + 1))
    period_cell, name_cell = f"$A{first}", f"B${top + 1}"
    if series:
        # one formula per bucket column: the column of Metrics is part of the formula text
        for offset, bucket in enumerate(series):
            column = data.Range(data.Cells(first, 2 + offset), data.Cells(last, 2 + offset))
            column.Formula2 = D.bucket_point(bucket, period_cell)
    else:
        block.Formula2 = D.team_point(spec.column, period_cell, name_cell)
    return {"top": top, "first": first, "last": last, "count": count}


def _add_chart(dash, spec, index: int, points: dict) -> None:
    data = dash.Parent.Worksheets("ChartData")
    row, column = D.chart_anchor(index)
    area = dash.Range(dash.Cells(row, column), dash.Cells(row + D.ROWS_PER_CHART - 1, column + D.COLUMNS_PER_CHART - 1))
    # AddChart2(Style, XlChartType, Left, Top, Width, Height, NewLayout): positional
    shape = dash.Shapes.AddChart2(-1, spec.chart_type, area.Left, area.Top, area.Width, area.Height, True)
    shape.Name = spec.name
    chart = shape.Chart
    while chart.SeriesCollection().Count:
        chart.SeriesCollection(1).Delete()
    labels = data.Range(data.Cells(points["first"], 1), data.Cells(points["last"], 1))
    for offset in range(points["count"]):
        column_cells = data.Range(data.Cells(points["first"], 2 + offset), data.Cells(points["last"], 2 + offset))
        series = chart.SeriesCollection().NewSeries()
        series.Name = "=" + _address(data, points["top"] + 1, 2 + offset)
        series.Values = column_cells
        series.XValues = labels
    chart.HasTitle = True
    chart.ChartTitle.Text = spec.title
    chart.HasLegend = True
    chart.Legend.Position = -4107                                   # xlLegendPositionBottom
    chart.DisplayBlanksAs = 1                                       # xlNotPlotted: gaps, not zeros
    if spec.axis_format:
        R.put_en_us(chart.Axes(2).TickLabels, "NumberFormat", spec.axis_format)
    note = dash.Cells(row + D.ROWS_PER_CHART, column)
    note.Formula2 = D.note_formula(_address(data, points["first"], 2, points["last"], 1 + points["count"]),
                                   spec.reason)


def export_charts(wb, folder) -> dict[str, str]:
    """Save every chart of the Dashboard as <name>.png in `folder`; returns {chart name: file}."""
    folder.mkdir(parents=True, exist_ok=True)
    dash, files = wb.Worksheets("Dashboard"), {}
    dash.Activate()
    for spec in D.CHARTS:
        target = folder / f"{spec.name}.png"
        chart_object = dash.ChartObjects(spec.name)
        chart_object.Activate()                 # a chart that is off the visible part of the sheet exports as an empty file
        if not chart_object.Chart.Export(str(target.resolve()), "PNG"):
            raise OSError(f"Excel could not export the chart {spec.name}")
        files[spec.name] = str(target)
    dash.Range("A1").Select()
    return files


def serial(day: dt.date | dt.datetime) -> float:
    return R.to_excel_serial(day)
