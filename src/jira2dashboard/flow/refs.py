"""Names shared by the formula modules (pure): the tables, the half-second margin and the "row of this table" index."""
HALF = "(0.5/86400)"
SI, IF_, IS = "StatusIntervals", "ItemFlow", "Issues"


def row_index(table: str) -> str:
    """The 1-based position of the current row in `table` (a table column that follows another table, row by row)."""
    return f"ROW()-ROW({table}[#Headers])"
