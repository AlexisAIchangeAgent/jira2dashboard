"""The Excel side of the flow sheets that needs COM (build time and tests).

At build time this module writes the plan of the flow sheets (FlowFormulas, ChartSpec) and the frame of the
Dashboard. After each extraction, VBA (modFlowBuild) builds the tables, formulas and charts from that plan: the
workbook needs no Python to refresh. Here are also the readers the tests use (tables, errors, PNG export).
"""

from jira2dashboard import sheet_names as S
from jira2dashboard.excel import com as R
from jira2dashboard.flow import blueprint as B
from jira2dashboard.flow import dashboard_spec as D
from jira2dashboard import manual as M
from jira2dashboard import error_codes as EC
from jira2dashboard import rovo_prompt as RP

EXCEL_ERRORS = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
                -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!", -2146826254: "#CALC!"}


def read_table(wb, sheet: str, name: str) -> tuple[list[str], list[list]]:
    """Header names and the raw values (Value2: dates are serial numbers, errors are negative integers)."""
    table = wb.Worksheets(S.of(sheet)).ListObjects(name)
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
ROVO_MACRO = "modActions.BtnRovoPrompt"
ROVO_LOAD_MACRO = "modActions.BtnLoadRovoCsv"


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
    ws = R.add_sheet_after_last(wb, RP.SPEC_SHEET)                  # the text behind the button Prompt for Rovo
    R.write_block(ws, 1, 1, [list(RP.SPEC_HEADER)] + [list(row) for row in RP.spec_rows()])
    ws.Columns("A:B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 120
    ws = R.add_sheet_after_last(wb, "ErrorCodes")                    # the support codes: what each one means and what to check
    R.write_block(ws, 1, 1, [list(EC.HEADER)] + [list(row) for row in EC.rows()])
    ws.Range("A1:E1").Font.Bold = True
    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B:E").ColumnWidth = 50
    write_dashboard_frame(wb, R.add_sheet_after_last(wb, "Dashboard") if dash is None else dash, button_macro)


def write_dashboard_frame(wb, dash, button_macro: str | None = None) -> None:
    """Title, the as-of date, the status cell (name dashStatus, its fields in the rows under it) and, with the button, the
    buttons Extract, Cancel, Read the Jira configuration and Rovo and the progress bar. Right after the Dashboard, the sheet
    Charts: its title and the team selector (B2, name dashTeam); the charts come after each extraction, on that sheet, so
    that they hide nothing of the Dashboard."""
    dash.Range("A1").Value = "Flow metrics"
    dash.Range("A1").Font.Bold = True
    dash.Range("E1").Value = "As of (UTC)"
    dash.Range("F1").Formula2 = "=IFERROR(AsOfUtc,\"\")"
    R.put_en_us(dash.Range("F1"), "NumberFormat", "yyyy-mm-dd hh:mm")
    dash.Range("A:R").ColumnWidth = 11
    dash.Range("A:A").ColumnWidth = 22                  # the names of the status rows
    dash.Range(D.STATUS_CELL).Value = "Ready"
    dash.Range(D.STATUS_CELL).Font.Bold = True
    wb.Names.Add(D.STATUS, "=" + S.of("Dashboard") + "!" + dash.Range(D.STATUS_CELL).Address)
    status = dash.Range(D.STATUS_CELL)
    dash.Range(dash.Cells(status.Row + 1, status.Column), dash.Cells(status.Row + D.STATUS_ROWS, status.Column)).Font.Bold = True
    charts = wb.Worksheets.Add(None, dash)              # positional: (Before, After)
    charts.Name = S.of(D.CHARTS_SHEET)
    charts.Range("A1").Value = "Flow metrics - charts"
    charts.Range("A1").Font.Bold = True
    charts.Range("A2").Value = "Team of the Distribution chart"
    charts.Range("A:R").ColumnWidth = 11
    charts.Range(D.SELECTOR_CELL).Value = D.ALL_TEAMS
    wb.Names.Add(D.SELECTOR, f"={S.of(D.CHARTS_SHEET)}!" + charts.Range(D.SELECTOR_CELL).Address)
    dash.Activate()
    if button_macro:
        R.add_progress_bar(dash, D.BAR_CELL, D.BUTTON_WIDTH, D.BAR_HEIGHT)
        for cell, caption, macro in ((D.ROVO_CELL, D.ROVO_CAPTION, ROVO_MACRO),
                                     (D.ROVO_LOAD_CELL, D.ROVO_LOAD_CAPTION, ROVO_LOAD_MACRO),
                                     (D.CONFIG_CELL, D.CONFIG_CAPTION, CONFIG_MACRO),
                                     (D.EXTRACT_CELL, D.EXTRACT_CAPTION, button_macro),
                                     (D.CANCEL_CELL, D.CANCEL_CAPTION, CANCEL_MACRO)):
            R.add_macro_button(dash, cell, caption, macro, D.BUTTON_WIDTH, D.BUTTON_HEIGHT)


def write_manual(wb, dash, mandatory: tuple[str, ...]) -> None:
    """The first steps on the Dashboard (from STEPS_CELL down, beside the charts: one line per row, no wrap, so the rows
    keep the height the charts are sized from) and the sheet Manual right after the Dashboard."""
    first = dash.Range(D.STEPS_CELL)
    R.write_block(dash, first.Row, first.Column, [[line] for line in M.first_steps()])
    first.Font.Bold = True
    ws = wb.Worksheets.Add(None, dash)                       # positional: (Before, After)
    ws.Name = S.of(M.MANUAL_SHEET)
    rows = M.manual(mandatory)
    R.write_block(ws, 1, 1, [[text] for _kind, text in rows])
    for n, (kind, _text) in enumerate(rows, start=1):
        if kind == M.HEADING:
            ws.Cells(n, 1).Font.Bold = True
    ws.Columns("A").ColumnWidth = 120


def export_charts(wb, folder) -> dict[str, str]:
    """Save every chart of the sheet Charts as <name>.png in `folder`; returns {chart name: file}."""
    folder.mkdir(parents=True, exist_ok=True)
    dash, files = wb.Worksheets(S.of(D.CHARTS_SHEET)), {}
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
