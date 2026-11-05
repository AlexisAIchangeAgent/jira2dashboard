"""The Excel side of the flow sheets that needs COM (build time and tests).

At build time this module writes the plan of the flow sheets (FlowFormulas, ChartSpec) and the frame of the
Dashboard. After each extraction, VBA (modFlowBuild) builds the tables, formulas and charts from that plan: the
workbook needs no Python to refresh. Here are also the readers the tests use (tables, errors, PNG export).
"""

from jira2dashboard.excel import com as R
from jira2dashboard.flow import blueprint as B
from jira2dashboard.flow import dashboard_spec as D
from jira2dashboard import manual as M

EXCEL_ERRORS = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
                -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!", -2146826254: "#CALC!"}


def read_table(wb, sheet: str, name: str) -> tuple[list[str], list[list]]:
    """Header names and the raw values (Value2: dates are serial numbers, errors are negative integers)."""
    table = wb.Worksheets(sheet).ListObjects(name)
    header = [h for h in table.HeaderRowRange.Value2[0]]
    body = table.DataBodyRange.Value2
    return header, [list(row) for row in body]


def error_cells(wb) -> dict[str, int]:
    """How many cells of each flow table show an Excel error, by column (a typo in a formula shows here)."""
    found: dict[str, int] = {}
    for name in ("StatusIntervals", "ItemFlow", "DailyLoad", "Metrics", "DataChecks", "IssueLinks", "FillRates"):
        header, rows = read_table(wb, name, name)
        for row in rows:
            for column, value in zip(header, row):
                if isinstance(value, int) and value in EXCEL_ERRORS:
                    key = f"{name}.{column} {EXCEL_ERRORS[value]}"
                    found[key] = found.get(key, 0) + 1
    return found


CANCEL_MACRO = "modActions.BtnCancelExtraction"
CONFIG_MACRO = "modActions.BtnReadJiraConfiguration"


def write_blueprint(wb, dash=None, button_macro: str | None = None) -> None:
    """The sheets Issues, StatusIntervals, IssueLinks (empty, raw data comes from the extraction), FlowFormulas, ChartSpec and
    ChecksSpec (read by modFlowBuild) and the frame of the Dashboard (`dash`: an existing sheet, or a new one)."""
    for name in ("Issues", "StatusIntervals", "IssueLinks"):
        R.add_sheet_after_last(wb, name)
    for name, header, rows in ((B.SHEET_FORMULAS, B.FORMULAS_HEADER, B.formula_rows()),
                               (B.SHEET_CHARTS, B.CHARTS_HEADER, B.chart_rows()),
                               (B.SHEET_CHECKS, B.CHECKS_HEADER, B.check_rows())):
        ws = R.add_sheet_after_last(wb, name)
        R.write_block(ws, 1, 1, [list(header)] + [list(row) for row in rows])      # text: a formula stays text here
        ws.Columns("A:B").ColumnWidth = 28
    write_dashboard_frame(wb, R.add_sheet_after_last(wb, "Dashboard") if dash is None else dash, button_macro)


def write_dashboard_frame(wb, dash, button_macro: str | None = None) -> None:
    """Title, the team selector (B2, name dashTeam), the as-of date, the status cell (name dashStatus) and, with the button, the
    buttons Extract, Cancel and Read the Jira configuration and the progress bar. The charts come after each extraction."""
    dash.Range("A1").Value = "Flow metrics"
    dash.Range("A1").Font.Bold = True
    dash.Range("A2").Value = "Team of the Distribution chart"
    dash.Range(D.SELECTOR_CELL).Value = D.ALL_TEAMS
    wb.Names.Add(D.SELECTOR, "=Dashboard!" + dash.Range(D.SELECTOR_CELL).Address)
    dash.Range("E1").Value = "As of (UTC)"
    dash.Range("F1").Formula2 = "=IFERROR(AsOfUtc,\"\")"
    R.put_en_us(dash.Range("F1"), "NumberFormat", "yyyy-mm-dd hh:mm")
    dash.Range("A:R").ColumnWidth = 11
    dash.Range(D.STATUS_CELL).Value = "Ready"
    wb.Names.Add(D.STATUS, "=Dashboard!" + dash.Range(D.STATUS_CELL).Address)
    if button_macro:
        R.add_macro_button(dash, "J1", D.EXTRACT_CAPTION, button_macro, 190, 26)
        R.add_macro_button(dash, D.CANCEL_CELL, D.CANCEL_CAPTION, CANCEL_MACRO, 80, 26)
        R.add_macro_button(dash, D.CONFIG_CELL, D.CONFIG_CAPTION, CONFIG_MACRO, 170, 26)
        R.add_progress_bar(dash, D.BAR_CELL, 190, 14)


def write_manual(wb, dash, mandatory: tuple[str, ...]) -> None:
    """The first steps on the Dashboard (from STEPS_CELL down, beside the charts: one line per row, no wrap, so the rows
    keep the height the charts are sized from) and the sheet Manual right after the Dashboard."""
    first = dash.Range(D.STEPS_CELL)
    R.write_block(dash, first.Row, first.Column, [[line] for line in M.first_steps()])
    first.Font.Bold = True
    ws = wb.Worksheets.Add(None, dash)                       # positional: (Before, After)
    ws.Name = M.MANUAL_SHEET
    rows = M.manual(mandatory)
    R.write_block(ws, 1, 1, [[text] for _kind, text in rows])
    for n, (kind, _text) in enumerate(rows, start=1):
        if kind == M.HEADING:
            ws.Cells(n, 1).Font.Bold = True
    ws.Columns("A").ColumnWidth = 120


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
