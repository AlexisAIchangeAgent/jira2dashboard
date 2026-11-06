"""Write the parameter sheets into a workbook (needs Excel)."""
from jira2dashboard.excel import com as R
from jira2dashboard.params.model import ParamSheet


# Iterations and PIs are typed at the client: an Example column (D, next to Period | Start | End) shows the form of each
# value. The readers (VBA ReadTable, standalone params_book) find columns by name and skip a row with no first cell,
# so these notes are never read as data.
EXAMPLES = {
    "Iterations": ("Period: PI#01-2026 It1 (the name of one iteration, the label of the charts)",
                   "Start: 2026-10-05 (a date, included)",
                   "End: 2026-10-26 (a date, excluded: 3 weeks later, the day the next iteration starts)"),
    "PIs": ("PI: PI#01-2026 (the value of the PI field of Jira)",
            "Start: 2026-10-05 (a date, included)",
            "End: 2027-02-08 (a date, excluded: a PI of 18 weeks holds 6 iterations of 3 weeks)"),
}


def write_params(wb, sheets: list[ParamSheet]) -> None:
    """One worksheet per parameter sheet, added after the last one: header row, then the rows."""
    for sheet in sheets:
        ws = R.add_sheet_after_last(wb, sheet.name)
        R.write_block(ws, 1, 1, [list(sheet.header)] + [list(row) for row in sheet.rows])
        if sheet.name in EXAMPLES:
            width = len(sheet.header) + 1
            R.write_block(ws, 1, width, [["Example"]] + [[text] for text in EXAMPLES[sheet.name]])
        del ws
