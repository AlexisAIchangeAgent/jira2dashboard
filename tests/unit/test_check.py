"""compute --check (standalone/check.py): the figures of a workbook calculated by Excel against the same figures made again by Python. Grids typed
here, as read_sheets gives them (row 1 the header, an empty cell None, a number a float)."""
import pytest

from jira2dashboard.flow import formulas as F
from jira2dashboard.standalone import __main__ as X
from jira2dashboard.standalone import check as C


def test_a_number_within_the_tolerance_an_empty_cell_and_the_empty_text_are_the_same():
    assert C.same(46323.5, 46323.5 + 5e-6) and not C.same(1.0, 1.001)                 # half a second of a date: the same moment
    assert C.same(None, "") and C.same("not computable", "not computable") and not C.same("A", "a")
    assert not C.same(True, 1.0) and not C.same(0.0, "")


EXCEL = [["Control", "Assertion", "Expected", "Found", "Helper"],
         ["Population", "Completeness", 28.0, 28.0, "x"],
         ["Velocity", "Accuracy", 19.0, 19.0, "y"],
         [None, None, None, None, None],
         ["Parameters in force", None, None, None, None]]                               # another block under the table
PYTHON = [["Control", "Assertion", "Expected", "Found"], ["Velocity", "Accuracy", 19, 19], ["Population", "Completeness", 28, 28]]


def test_the_rows_are_joined_on_their_keys_and_the_columns_both_hold_are_compared():
    got = C.compare_table("Reconciliation", ("Control",), EXCEL, PYTHON)
    assert got.ok and (got.rows, got.cells) == (2, 6)                                   # the Helper column and the block below are not compared


def test_a_cell_that_differs_and_a_row_on_one_side_are_listed():
    python = [r[:] for r in PYTHON] + [["Unique", "Occurrence", 0, 0]]
    python[1][3] = 18
    got = C.compare_table("Reconciliation", ("Control",), EXCEL, python)
    assert got.differences == [C.Difference("Velocity", "Found", 19.0, 18)]
    assert (got.only_excel, got.only_python) == ([], ["Unique"]) and not got.ok


def test_the_lists_of_metricitems_are_compared_as_sets_of_rows_whatever_their_order():
    width = F.metricitems_column(3, F.METRICITEMS_BLOCKS[3][-1][0])
    names = [""] * width
    for block in range(4):
        for name, _formula in F.METRICITEMS_BLOCKS[block]:
            names[F.metricitems_column(block, name) - 1] = name
    def row(key, block=0):
        cells = [None] * width
        cells[F.metricitems_column(block, "Key") - 1] = key
        return cells
    excel = [["Done in an iteration: ..."] + [None] * (width - 1), names, row("A-1"), row("B-2"), row("C-3", 2)]
    python = [names, row("B-2"), row("A-1"), row("C-3", 2)]
    assert C.compare_lists(excel, python).ok
    got = C.compare_lists(excel, python[:-1])
    assert len(got.only_excel) == 1 and got.only_excel[0].startswith("Cut-off") and "C-3" in got.only_excel[0] and not got.only_python
    got = C.compare_lists(excel[:-1], python)
    assert (got.only_excel, len(got.only_python)) == ([], 1) and "C-3" in got.only_python[0]


def test_a_missing_sheet_and_dailyload_are_said_and_the_summary_names_what_to_look_at():
    checks = C.check({"Reconciliation": EXCEL}, {"Reconciliation": PYTHON})
    by_table = {c.table: c for c in checks}
    assert by_table["Reconciliation"].ok and by_table["DailyLoad"].note.startswith("not compared")
    assert by_table["Metrics"].note == "the sheet Metrics_from_excel is missing in the workbook" and not by_table["Metrics"].ok      # J2D-FLOW-012
    assert "0 differ" in C.summary(checks) and "to look at: Metrics" in C.summary(checks)


def test_the_page_shows_the_tables_that_differ_first_and_escapes_what_it_shows():
    python = [r[:] for r in PYTHON]
    python[1][1] = "<b>"
    page = C.page([C.compare_table("Reconciliation", ("Control",), EXCEL, python), C.TableCheck("ItemFlow", 3, 9)], "a.xlsm", "b.xlsx")
    assert page.index("Reconciliation: DIFF") < page.index("ItemFlow: OK") and "&lt;b&gt;" in page and "<b>" not in page
    assert "<script" not in page and "http" not in page


def test_check_needs_a_workbook_that_excel_calculated():
    with pytest.raises(SystemExit) as stop:
        X.parse(["--workbook", "results.xlsx", "--check"], "compute")
    assert stop.value.code == 2
    assert X.parse(["--workbook", "book.xlsm", "--check"], "compute").check
    assert not hasattr(X.parse(["--workbook", "book.xlsm"]), "check") or not X.parse(["--workbook", "book.xlsm"]).check


def test_every_table_the_python_path_writes_is_compared_or_says_why_not():
    tables = set(F.TABLES) | {"StatusRules"}
    assert tables <= set(C.TABLE_KEYS) | set(C.NOT_COMPARED), sorted(tables - set(C.TABLE_KEYS) - set(C.NOT_COMPARED))
    assert all(set(keys) <= set(F.TABLES.get(table, keys)) for table, keys in C.TABLE_KEYS.items())          # the keys are columns of the table


def test_compute_never_writes_over_the_workbook_it_reads():
    with pytest.raises(SystemExit) as stop:
        X.parse(["--workbook", "C:/x/jira2dashboard-compute.xlsx"], "compute")
    assert stop.value.code == 2
    assert X.parse(["--workbook", "C:/x/jira2dashboard-compute.xlsx", "--out", "C:/x/again.xlsx"], "compute").out
