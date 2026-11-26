"""The type of every column of the raw sheets written from Jira, and the check of the first filled value of each column.

Why: Jira's schema is no guarantee (the OpenAPI says `created` of a changelog is a `date-time`, the real answer is a number of
milliseconds: v0.7.0 wrote an empty date column and nobody saw it). After an extraction the workbook reads the CELLS it wrote:
the first filled value of each column must be of the type the column's label promises (a `...Utc` column holds a datetime).
A problem never stops the run: it is a WARN line of the log with a support code per data type, and a field of the final status.

The spec (`RAW_TYPES`) is data: written at build into the sheet RawTypesSpec, read by `modHealth.TypeProblems` (VBA); the Python
path calls `check` on its rows. The two sides must agree (tests/unit/test_raw_types.py, and the Excel scenarios).

Values seen by `kind_of`: empty (None, ""), text, number, date, boolean, error (a VBA error cell), other (a list or an object).
A log line quotes names, types, counts and the SHAPE of the value (digits as 9, letters as A or a), never the value.
"""
from __future__ import annotations

import datetime as dt
import re
from collections import Counter
from typing import Iterable, NamedTuple, Sequence

# data type -> support code (error_codes.py). A code is never reused.
TYPE_CODES = {
    "datetime": "J2D-DATA-004", "date": "J2D-DATA-005", "id": "J2D-DATA-006", "key": "J2D-DATA-007",
    "integer": "J2D-DATA-008", "number": "J2D-DATA-009", "boolean": "J2D-DATA-010", "enum": "J2D-DATA-011",
    "list": "J2D-DATA-012", "text": "J2D-DATA-013",
}
SAMPLE = 20                                           # the kinds of the first filled values are counted for the log
FILL_LINE_MAX = 270                                   # a log line is cut at 300 characters (modLog): the fill counts go on a next line
DATE_FIRST, DATE_LAST = dt.datetime(2000, 1, 1), dt.datetime(2100, 1, 1)
KEY_FORMAT = re.compile(r"^[A-Z][A-Z0-9_]*-[0-9]+$")  # a ticket key: project key, dash, number
SPRINT_STATES = ("future", "active", "closed")        # Jira Software: the state of a sprint
STATUS_CATEGORIES = ("new", "indeterminate", "done", "undefined")   # Jira: the key of a status category ("undefined": No Category)
FOUND_ABSENT = "missing"                              # a column the sheet does not have


class Rule(NamedTuple):
    table: str
    column: str
    type: str
    required: bool                                    # all cells empty is a problem too (the v0.7.0 failure)
    source: str                                       # where Jira says it comes from, for the log only
    allowed: tuple = ()                               # the closed list of an enum


def _rules(table: str, rows: Iterable[tuple]) -> list[Rule]:
    return [Rule(table, *row) for row in rows]


ISSUES_RULES = _rules("Issues", [
    ("Key", "key", True, "key"),
    ("Id", "id", True, "id"),
    ("Project", "text", True, "fields.project.key"),
    ("Type", "text", True, "fields.issuetype.name"),
    ("WorkType", "text", False, "fields.<WorkTypeField>"),
    ("CreatedUtc", "datetime", True, "fields.created"),
    ("Status", "text", True, "fields.status.name"),
    ("Parent", "key", False, "fields.parent.key"),
    ("StatusCategory", "enum", False, "fields.status.statusCategory.key", STATUS_CATEGORIES),
    ("ResolvedUtc", "datetime", False, "fields.resolutiondate"),
    ("FixVersions", "list", False, "fields.fixVersions[].name"),
    ("StoryPoints", "number", False, "fields.<StoryPointsField>"),
    ("PI", "text", False, "fields.<PIField>"),
    ("BV", "number", False, "fields.<BusinessValueField>"),
    ("TeamValue", "text", False, "fields.<TeamField>"),
    ("SprintState", "enum", False, "fields.<sprint>[].state", SPRINT_STATES),
    ("SprintStart", "datetime", False, "fields.<sprint>[].startDate"),
    ("SprintEnd", "datetime", False, "fields.<sprint>[].endDate"),
    ("SprintCount", "integer", True, "fields.<sprint>"),
    ("TargetStart", "date", False, "fields.<TargetStartField>"),
    ("TargetEnd", "date", False, "fields.<TargetEndField>"),
    ("KeyChanges", "integer", True, "computed: changelog items of the field Key"),
])
INTERVALS_RULES = _rules("StatusIntervals", [
    ("Key", "key", True, "computed"),
    ("Workflow", "text", True, "computed"),
    ("Seq", "integer", True, "computed"),
    ("Status", "text", True, "computed: histories[].items[].toString"),
    ("Category", "text", True, "computed: sheet Workflows"),
    ("EnteredUtc", "datetime", True, "computed: histories[].created"),
    ("LeftUtc", "datetime", False, "computed: histories[].created"),
    ("DurationDays", "number", True, "computed"),
    ("IsOpen", "boolean", True, "computed"),
    ("ChangeId", "id", False, "histories[].id"),
    ("Check", "text", False, "computed"),
])
LINKS_RULES = _rules("IssueLinks", [
    ("Blocker", "key", False, "fields.issuelinks[].outwardIssue.key"),
    ("Blocked", "key", False, "key"),
])
CHANGELOG_RULES = _rules("Changelog", [
    ("Key", "key", True, "key"),
    ("IssueId", "id", True, "issueChangeLogs[].issueId"),
    ("ChangeId", "id", True, "histories[].id"),
    ("CreatedUtc", "datetime", True, "histories[].created"),
    ("CreatedText", "text", True, "histories[].created"),
    ("FieldId", "text", True, "histories[].items[].fieldId"),
    ("Field", "text", True, "histories[].items[].field"),
    ("FieldType", "text", False, "histories[].items[].fieldtype"),
    ("From", "text", False, "histories[].items[].from"),
    ("FromString", "text", False, "histories[].items[].fromString"),
    ("To", "text", False, "histories[].items[].to"),
    ("ToString", "text", False, "histories[].items[].toString"),
])
PI_INTERVALS_RULES = _rules("PIIntervals", [          # no PIField, no stretch: every column may be empty
    ("Key", "key", False, "computed"),
    ("PI", "text", False, "computed: histories[].items[].toString"),
    ("FromUtc", "datetime", False, "computed: Issues.CreatedUtc or histories[].created"),
    ("ToUtc", "datetime", False, "computed: histories[].created"),
    ("IsOpen", "boolean", False, "computed"),
])
RAW_TYPES: tuple[Rule, ...] = tuple(ISSUES_RULES + INTERVALS_RULES + LINKS_RULES + CHANGELOG_RULES + PI_INTERVALS_RULES)
SPEC_SHEET = "RawTypesSpec"
SPEC_HEADER = ("Table", "Column", "Type", "Code", "Required", "Source", "Allowed")


def spec_rows() -> list[tuple]:
    """The rows of the sheet RawTypesSpec (SPEC_HEADER): what modHealth reads."""
    return [(r.table, r.column, r.type, TYPE_CODES[r.type], "yes" if r.required else "no", r.source, ", ".join(r.allowed))
            for r in RAW_TYPES]                                         # yes/no, not a Boolean: CStr(True) is "Vrai" on a French PC


# --- what a value is -----------------------------------------------------------------------------------------------

def kind_of(value) -> str:
    """empty | text | number | date | boolean | error | other: the same words as modHealth.KindOf on a cell."""
    if value is None or (isinstance(value, str) and value == ""):
        return "empty"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, (dt.datetime, dt.date)):
        return "date"
    if isinstance(value, str):
        return "text"                                              # a Python str is never an Excel error (the VBA side reads IsError)
    return "other"


def _as_datetime(value) -> dt.datetime:
    return value if isinstance(value, dt.datetime) else dt.datetime(value.year, value.month, value.day)


def fits(type_: str, value, allowed: Sequence[str] = ()) -> bool:
    """Is this (filled) value of the type? The rules of the table of raw_types.py's docstring; modHealth.Fits is the VBA side."""
    kind = kind_of(value)
    if type_ in ("datetime", "date"):
        if kind != "date":
            return False
        moment = _as_datetime(value)
        return DATE_FIRST <= moment < DATE_LAST and (type_ == "datetime" or moment.time() == dt.time(0))
    if type_ == "id":
        return kind == "text" and value.isascii() and value.isdigit()
    if type_ == "key":
        return kind == "text" and KEY_FORMAT.fullmatch(value) is not None
    if type_ == "integer":
        return kind == "number" and float(value).is_integer()
    if type_ == "number":
        return kind == "number"
    if type_ == "boolean":
        return kind == "boolean"
    if type_ == "enum":
        return kind == "text" and value in allowed
    if type_ == "list":
        return kind == "text" and all(part.strip() for part in value.split(","))
    return kind == "text"                                          # text


def shape_of(value) -> str:
    """The type and the shape of a value, never the value: "Double:9999999999999", "String:AAA-99". 40 characters at most.
    The type words are those of modFlow.DateShape (VBA TypeName): Double, Date, String, Boolean, Empty (and Error, VBA only:
    a Python value is never an Excel error)."""
    kind = kind_of(value)
    if kind in ("empty", "other"):
        return "Empty" if kind == "empty" else "Object"
    if kind == "boolean":
        return "Boolean"
    if kind == "date":
        text, word = _as_datetime(value).strftime("%Y-%m-%dT%H:%M:%S"), "Date"
    elif kind == "number":
        text, word = (str(int(value)) if float(value).is_integer() else format(float(value), ".15g")), "Double"      # as modHealth.PlainNumber
    else:
        text, word = value, "String"
    masked = "".join("9" if c.isdigit() and c.isascii() else "A" if "A" <= c <= "Z" else "a" if "a" <= c <= "z"
                     else c if " " <= c < "\x7f" else "?" for c in text[:40])
    return f"{word}:{masked}{'...' if len(text) > 40 else ''}"


def digits_of(value) -> int:
    """How many digits a number or a digit text has (13 = epoch milliseconds, 10 = epoch seconds); 0 for anything else."""
    kind = kind_of(value)
    if kind == "number" and float(value).is_integer():
        return len(str(abs(int(value))))
    if kind == "text" and value.isascii() and value.isdigit():
        return len(value)
    return 0


def hint(rule: Rule, found: str, first=None) -> str:
    """What to look at, in one sentence: a fixed text per (expected type, found kind). No value in it."""
    expected = rule.type
    if found == FOUND_ABSENT:
        return "the column is not in the sheet: the sheet and RawTypesSpec differ (an old workbook?)"
    if found == "empty":
        if expected in ("datetime", "date"):
            return ("no cell holds a date: Jira's date was not read (see WARN J2D-DATA-003 and the INFO line 'changelog date key')"
                    if rule.required else "no ticket has this date: fine when the field is not used")
        return ("a required column is empty: the extraction wrote nothing there (check the field and its id in Config_from_user)"
                if rule.required else "no ticket has a value: fine when the field is not configured or not used")
    if expected in ("datetime", "date"):
        digits = digits_of(first)
        if digits >= 12:
            return f"{digits} digits: epoch milliseconds, Jira sent a number and the date was not converted (modText.ParseJiraDate)"
        if digits >= 9:
            return f"{digits} digits: epoch seconds, Jira sent a number and the date was not converted (modText.ParseJiraDate)"
        if found == "date":
            return "a date outside 2000-2099, or with a time of day where Jira sends a day: a wrong unit or a typo in Jira"
        if found == "text":
            return "a text that is not read as a date: Jira's format changed (see the shape; ParseJiraDate reads ISO and epoch)"
        return "not a date: the cell holds another type than a date"
    if expected in ("id", "key"):
        if found == "number":
            return "Excel turned the text into a number (leading apostrophe lost) or Jira sent a number: keep ids as text"
        return "not a Jira id or key (ABC-12 or digits): the field is not the one expected"
    if expected in ("integer", "number"):
        if found == "text":
            return "a text where a number is expected: Jira sent text, or the custom field is of another type (check its id in Config_from_user)"
        if found == "number":
            return "a number with decimals where a whole number is expected"
        return "not a number: the cell holds another type"
    if expected == "boolean":
        return "not TRUE/FALSE: the cell holds another type"
    if expected == "enum":
        return f"not one of {', '.join(rule.allowed)}: Jira added a value or the field is another one"
    if expected == "list":
        return "an item of the list is empty: a separator without a value"
    return "a text expected: Excel converted it (apostrophe lost) or Jira typed the value differently"      # text


# --- the check ------------------------------------------------------------------------------------------------------

def column_letter(number: int) -> str:
    """1 -> A, 27 -> AA."""
    letters = ""
    while number > 0:
        number, rest = divmod(number - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


class TypeProblem(NamedTuple):
    table: str
    column: str
    type: str                       # expected
    code: str
    found: str                      # kind of the first filled value, "empty" or "missing"
    source: str
    cell: str                       # first filled cell, "" when none
    rows: int
    filled: int
    kinds: str                      # "number x 20": the kinds of the first filled values
    shape: str                      # shape of the first filled value, never the value
    hint: str

    @property
    def name(self) -> str:
        return f"{self.table}.{self.column}"

    def lines(self) -> list[tuple[str, str]]:
        """(level, message) of the log: a WARN line with the code, the column, the types and where Jira says it comes from, then two INFO
        lines with the same code and column (the counts and the shape, the hint). Three lines because the VBA log cuts a line at 300
        characters: one line per problem lost its hint, the part that helps most. Names, counts and a shape, never a value."""
        head = f"{self.code} | {self.name}"
        return [("WARN", f"{head} | expected {self.type} | found {self.found} | source {self.source} | first filled {self.cell or '-'}"),
                ("INFO", f"{head} | rows {self.rows}, filled {self.filled}, empty {self.rows - self.filled} | "
                         f"kinds in first {SAMPLE}: {self.kinds or '-'} | value shape {self.shape}"),
                ("INFO", f"{head} | hint: {self.hint}")]


class TypeReport(NamedTuple):
    problems: list[TypeProblem]
    tables: int                     # tables read (with at least one row)
    columns: int                    # columns of the spec in those tables
    checked: int                    # columns that had a filled cell
    fills: dict[str, list[tuple[str, int]]]   # table -> [(column, filled cells)], in sheet order
    rows: dict[str, int]


def _is_empty(value) -> bool:
    return kind_of(value) == "empty"


def check(tables: dict[str, tuple[Sequence[str], Sequence[Sequence]]]) -> TypeReport:
    """The problems of {table: (header, rows)}: for each column of RAW_TYPES, the first filled value must fit the type; a required
    column with no filled cell is a problem too; an optional one stays silent. A table without a row is not read."""
    problems: list[TypeProblem] = []
    fills: dict[str, list[tuple[str, int]]] = {}
    rows_of: dict[str, int] = {}
    seen_columns = checked = 0
    for table in dict.fromkeys(r.table for r in RAW_TYPES):
        if table not in tables or not tables[table][1]:
            continue
        header, rows = tables[table]
        rows_of[table] = len(rows)
        fills[table] = []
        for rule in (r for r in RAW_TYPES if r.table == table):
            seen_columns += 1
            if rule.column not in header:
                problems.append(TypeProblem(table, rule.column, rule.type, TYPE_CODES[rule.type], FOUND_ABSENT, rule.source, "",
                                            len(rows), 0, "", "Empty", hint(rule, FOUND_ABSENT)))
                continue
            at = header.index(rule.column)
            filled = [(n, row[at]) for n, row in enumerate(rows) if at < len(row) and not _is_empty(row[at])]
            fills[table].append((rule.column, len(filled)))
            if not filled:
                if rule.required:
                    problems.append(TypeProblem(table, rule.column, rule.type, TYPE_CODES[rule.type], "empty", rule.source, "",
                                                len(rows), 0, "", "Empty", hint(rule, "empty")))
                continue
            checked += 1
            position, first = filled[0]
            if fits(rule.type, first, rule.allowed):
                continue
            found = kind_of(first)
            kinds = Counter(kind_of(v) for _n, v in filled[:SAMPLE])
            problems.append(TypeProblem(
                table, rule.column, rule.type, TYPE_CODES[rule.type], found, rule.source,
                f"{column_letter(at + 1)}{position + 2}", len(rows), len(filled),
                ", ".join(f"{k} x {n}" for k, n in sorted(kinds.items())), shape_of(first), hint(rule, found, first)))
    return TypeReport(problems, len(rows_of), seen_columns, checked, fills, rows_of)


def summary_line(report: TypeReport) -> str:
    return (f"column types | tables {report.tables} | columns {report.columns} | checked {report.checked} | "
            f"problems {len(report.problems)}")


def table_fill_lines(report: TypeReport, table: str) -> list[str]:
    """The filled cells of each column of a table, so that an empty optional column shows too: one line, continued on a next line
    (`types T | filled (continued): ...`) when it would pass FILL_LINE_MAX characters (modHealth.TypeProblems does the same)."""
    columns = report.fills.get(table) or []
    if not columns:
        return []
    lines, line, count = [], f"types {table} | rows {report.rows[table]} | filled: ", 0
    for column, filled in columns:
        item = f"{column} {filled}"
        if count and len(line) + 2 + len(item) > FILL_LINE_MAX:
            lines.append(line)
            line, count = f"types {table} | filled (continued): ", 0
        line += (", " if count else "") + item
        count += 1
    return lines + [line]


def fill_lines(report: TypeReport) -> list[str]:
    return [line for table in report.fills for line in table_fill_lines(report, table)]


def log_lines(report: TypeReport) -> list[tuple[str, str]]:
    """(level, message) of the whole check, in the order modHealth.TypeProblems writes them: per table the three lines of each of its
    problems, then its fill lines; at the end the summary."""
    out: list[tuple[str, str]] = []
    for table in report.fills:
        for problem in report.problems:
            if problem.table == table:
                out += problem.lines()
        out += [("INFO", line) for line in table_fill_lines(report, table)]
    return out + [("INFO", summary_line(report))]


def status_field(report: TypeReport, shown: int = 5) -> str:
    """` | Column types: n column(s), first Changelog.CreatedUtc expected datetime found empty; ...` or "" when all is well."""
    if not report.problems:
        return ""
    names = "; ".join(f"{p.name} expected {p.type} found {p.found}" for p in report.problems[:shown])
    more = "; ..." if len(report.problems) > shown else ""
    return f" | Column types: {len(report.problems)} column(s), first {names}{more}"
