"""`run.ps1 compute --check` (step F): the figures that Excel calculated in a workbook against the same figures made again by Python.

The workbook is an .xlsm saved by Excel after its buttons ran (the formulas' last values are kept in the file); `compute` makes the results again
from its raw sheets and writes them as values; this module reads both files, joins the rows of each result table on its keys and compares every
cell of the columns both hold: numbers within TOLERANCE (a date is a number of days: half a second is about 6e-6), texts exactly, an empty cell =
the empty text. The lists of MetricItems are compared as sets of rows, block by block. A known difference of the Python path (DailyLoad stops at
AsOf) is not compared and says so. Pure but for the file reads; the page is plain HTML (no script, no font from outside).
"""
import html
from dataclasses import dataclass, field

from jira2dashboard import sheet_names as S
from jira2dashboard.flow import formulas as F

TOLERANCE = 1e-5
SHOWN = 200                                     # differences listed per table on the page (the count says the rest)
TABLE_KEYS = {
    "Metrics": F.METRICS_KEYS, "PIMetrics": ("PI",), "PIStatusTime": ("PI", "Status"), "VelocityCheck": F.METRICS_KEYS,
    "FlowTimeCheck": F.METRICS_KEYS, "FlowEfficiencyCheck": F.METRICS_KEYS, "FlowLoadCheck": F.METRICS_KEYS, "FlowDistCheck": F.METRICS_KEYS,
    "IterationFeatures": F.ITERATIONFEATURES_KEYS, "ItemFlow": ("Key",), "Reconciliation": ("Control",),
}
NOT_COMPARED = {"DailyLoad": "the Python path stops at AsOf, the workbook at the end of the last iteration"}
LISTS = "MetricItems"


@dataclass(frozen=True)
class Difference:
    row: str                       # the keys of the row, joined by " | "
    column: str
    excel: object
    python: object


@dataclass
class TableCheck:
    table: str
    rows: int = 0                  # rows found on both sides
    cells: int = 0                 # cells compared
    differences: list[Difference] = field(default_factory=list)
    only_excel: list[str] = field(default_factory=list)       # keys of the rows Python does not have
    only_python: list[str] = field(default_factory=list)
    note: str = ""

    @property
    def ok(self) -> bool:
        return not (self.differences or self.only_excel or self.only_python)


def _blank(value):
    return "" if value is None else value


def same(excel, python) -> bool:
    excel, python = _blank(excel), _blank(python)
    if isinstance(excel, bool) or isinstance(python, bool):                 # TRUE is no 1: a boolean column that became a number shows
        return type(excel) is type(python) and excel == python
    if isinstance(excel, (int, float)) and isinstance(python, (int, float)):
        return abs(excel - python) <= TOLERANCE
    return excel == python


def _cell(row: list, at: int):
    return row[at] if at < len(row) else None


def _rows(grid: list[list], keys: tuple[str, ...]) -> tuple[list[str], dict[str, list]]:
    """The header and the rows by their keys. The table ends at its first row with no key: the one blank row of an empty table, the blank row
    before another block below it (the parameters in force under the table Reconciliation)."""
    header = ["" if c is None else str(c) for c in (grid[0] if grid else [])]
    at = [header.index(k) for k in keys if k in header]
    if len(at) < len(keys):
        return header, {}
    rows = {}
    for row in grid[1:]:
        values = [_blank(_cell(row, i)) for i in at]
        if all(v == "" for v in values):
            break
        rows[" | ".join(_text(v) for v in values)] = row
    return header, rows


def _text(value) -> str:
    value = _blank(value)
    if isinstance(value, float):
        return f"{value:.6g}"
    return str(value)


def compare_table(table: str, keys: tuple[str, ...], excel: list[list], python: list[list]) -> TableCheck:
    """Every cell of the columns that both sheets hold (the workbook has more: helper columns, blocks beside the table), row by row on the keys."""
    result = TableCheck(table)
    excel_header, excel_rows = _rows(excel, keys)
    python_header, python_rows = _rows(python, keys)
    if not python_rows and not excel_rows:
        result.note = "no row on either side"
        return result
    columns = [c for c in python_header if c and c in excel_header and c not in keys]
    result.only_excel = [k for k in excel_rows if k not in python_rows]
    result.only_python = [k for k in python_rows if k not in excel_rows]
    for key, row in python_rows.items():
        if key not in excel_rows:
            continue
        result.rows += 1
        for column in columns:
            a, b = _cell(excel_rows[key], excel_header.index(column)), _cell(row, python_header.index(column))
            result.cells += 1
            if not same(a, b):
                result.differences.append(Difference(key, column, _blank(a), _blank(b)))
    return result


def _blocks(grid: list[list], header_row: int) -> list[set[tuple]]:
    """The lists of MetricItems as sets of rows, from the row of the column names (the workbook has a title row above it, the Python path does not)."""
    out = []
    for block in range(len(F.METRICITEMS_BLOCKS)):
        first = F.metricitems_column(block, F.METRICITEMS_BLOCKS[block][0][0]) - 1
        width = len(F.METRICITEMS_BLOCKS[block])
        rows = set()
        for row in grid[header_row + 1:]:
            values = tuple(_text(_cell(row, first + i)) for i in range(width))
            if any(v != "" for v in values):
                rows.add(values)
        out.append(rows)
    return out


def _names_row(grid: list[list]) -> int:
    first = F.METRICITEMS_BLOCKS[0][0][0]
    return next((n for n, row in enumerate(grid[:3]) if row and row[0] == first), 0)


def compare_lists(excel: list[list], python: list[list]) -> TableCheck:
    """The four lists of MetricItems: the same rows on both sides (as texts: a number to 6 significant digits), whatever their order."""
    result = TableCheck(LISTS)
    for block, (a, b) in enumerate(zip(_blocks(excel, _names_row(excel)), _blocks(python, _names_row(python)))):
        name = F.METRICITEMS_TITLES[block].split(":")[0]
        result.rows += len(a & b)
        result.cells += len(a | b)
        result.only_excel += [f"{name}: {' | '.join(r)}" for r in sorted(a - b)]
        result.only_python += [f"{name}: {' | '.join(r)}" for r in sorted(b - a)]
    return result


def check(excel: dict[str, list[list]], python: dict[str, list[list]]) -> list[TableCheck]:
    """`excel`, `python`: the grids by LOGICAL sheet name. A table missing on a side is a check with a note."""
    out = []
    for table, keys in TABLE_KEYS.items():
        if table not in excel or table not in python:
            out.append(TableCheck(table, note=f"the sheet {S.of(table)} is missing in the {'workbook' if table not in excel else 'Python results'}"))
            continue
        out.append(compare_table(table, keys, excel[table], python[table]))
    if LISTS in excel and LISTS in python:
        out.append(compare_lists(excel[LISTS], python[LISTS]))
    out += [TableCheck(table, note=f"not compared: {why}") for table, why in NOT_COMPARED.items()]
    return out


def summary(checks: list[TableCheck]) -> str:
    wrong = [c.table for c in checks if not c.ok or c.note.startswith("the sheet")]
    cells = sum(c.cells for c in checks)
    return (f"{len(checks)} tables, {cells} cells, {sum(len(c.differences) for c in checks)} differ, "
            f"{sum(len(c.only_excel) + len(c.only_python) for c in checks)} rows on one side only"
            + (f" | to look at: {', '.join(wrong)}" if wrong else ""))


def page(checks: list[TableCheck], workbook: str, results: str) -> str:
    """The page expected (Excel, the workbook) / obtained (Python, compute): one section per table, those that differ open and first."""
    e = html.escape
    ordered = sorted(checks, key=lambda c: c.ok)
    parts = [f"<h1>compute --check</h1><p>Workbook (Excel): {e(workbook)}<br>Python results: {e(results)}</p>",
             f'<p class="{"ok" if all(c.ok for c in checks) else "diff"}">{e(summary(checks))}</p>']
    for c in ordered:
        state = "OK" if c.ok else "DIFF"
        parts.append(f'<details{"" if c.ok else " open"}><summary>{e(c.table)}: {state} - {c.rows} rows, {c.cells} cells'
                     f'{", " + e(c.note) if c.note else ""}</summary>')
        if c.differences:
            parts.append("<table><tr><th>Row</th><th>Column</th><th>Excel</th><th>Python</th></tr>")
            parts += [f"<tr><td>{e(d.row)}</td><td>{e(d.column)}</td><td>{e(_text(d.excel))}</td><td>{e(_text(d.python))}</td></tr>"
                      for d in c.differences[:SHOWN]]
            parts.append("</table>")
            if len(c.differences) > SHOWN:
                parts.append(f"<p>... and {len(c.differences) - SHOWN} more</p>")
        for side, keys in (("only in the workbook", c.only_excel), ("only in the Python results", c.only_python)):
            if keys:
                parts.append(f"<p>Rows {side}: {len(keys)}</p><ul>" + "".join(f"<li>{e(k)}</li>" for k in keys[:SHOWN]) + "</ul>")
        parts.append("</details>")
    style = ("body{font-family:Segoe UI,Arial,sans-serif;margin:16px}table{border-collapse:collapse}td,th{border:1px solid #bbb;padding:2px 6px}"
             ".ok{color:#1a7f37}.diff{color:#b42318;font-weight:bold}summary{cursor:pointer;margin:6px 0}")
    return (f'<!doctype html><html lang="en"><head><meta charset="utf-8"><title>compute --check</title><style>{style}</style></head>'
            f'<body>{"".join(parts)}</body></html>')
