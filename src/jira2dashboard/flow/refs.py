"""Names shared by the formula modules (pure): the tables, the half-second margin and the "row of this table" index."""
from jira2dashboard import sheet_names as S

HALF = "(0.5/86400)"
SI, IF_, IS = "StatusIntervals", "ItemFlow", "Issues"
TRANSITIONS_LAST, PIS_LAST = "cfg_TransitionsLast", "cfg_PIsLast"       # control cells: the last used row of a parameter sheet


def row_index(table: str) -> str:
    """The 1-based position of the current row in `table` (a table column that follows another table, row by row)."""
    return f"ROW()-ROW({table}[#Headers])"


def bounded(sheet: str, column: str, last: str) -> str:
    """A column of a parameter sheet from row 2 to its last used row (the control cell `last`), never the whole column.

    COUNTIFS on a whole column scans its 1 048 576 rows: 11 s for the 7 284 rows of StatusIntervals against a sheet of
    27 rows, and 0.1 s with the range cut at the last row (SUMIFS and MATCH stop at the used range, COUNTIFS does not).
    """
    sheet = S.of(sheet)                                  # the real name of the sheet: Transitions -> Transitions_from_user
    return f"{sheet}!${column}$2:INDEX({sheet}!${column}:${column},{last})"


def last_row_cell(sheet: str) -> str:
    """The formula of a control cell: the last row with a text in column A of `sheet` (2 when the sheet has no row)."""
    return f'=MAX(2,IFERROR(MATCH(REPT("z",255),{S.of(sheet)}!$A:$A),2))'
