"""Write the parameter sheets into a workbook (needs Excel)."""
from jira2dashboard.excel import com as R
from jira2dashboard.params.model import ParamSheet


def write_params(wb, sheets: list[ParamSheet]) -> None:
    """One worksheet per parameter sheet, added after the last one: header row, then the rows."""
    for sheet in sheets:
        ws = R.add_sheet_after_last(wb, sheet.name)
        R.write_block(ws, 1, 1, [list(sheet.header)] + [list(row) for row in sheet.rows])
        del ws
