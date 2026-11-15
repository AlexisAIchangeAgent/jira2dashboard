"""The type check of the raw sheets (flow/raw_types.py): the spec, the type of a value, the first filled value of a column,
the log line (names, counts and shapes, never a value), and one support code per data type."""
import datetime as dt
import re
from pathlib import Path

import pytest

from jira2dashboard import error_codes as E
from jira2dashboard.flow import formulas as F
from jira2dashboard.flow import raw_types as T

D = dt.datetime(2026, 11, 14, 9, 30, 0)


# --- the spec -----------------------------------------------------------------------------------------------------------------

def test_every_raw_column_has_a_rule_in_sheet_order():
    """A new column of a raw sheet forces a decision about its type."""
    wanted = {"Issues": F.ISSUES, "StatusIntervals": F.INTERVALS_RAW, "IssueLinks": ("Blocker", "Blocked"), "Changelog": F.CHANGELOG_RAW}
    for table, columns in wanted.items():
        assert tuple(r.column for r in T.RAW_TYPES if r.table == table) == tuple(columns), table


def test_every_column_called_utc_is_a_datetime():
    assert all(r.type == "datetime" for r in T.RAW_TYPES if r.column.endswith("Utc"))


def test_every_type_has_its_own_code_known_to_the_support_table():
    assert set(T.TYPE_CODES) == {r.type for r in T.RAW_TYPES}                          # every type is used, none without a code
    codes = list(T.TYPE_CODES.values())
    assert len(set(codes)) == len(codes) == 10 and all(code in E.BY_CODE for code in codes)
    assert codes == [f"J2D-DATA-{n:03d}" for n in range(4, 14)]                       # the numbers are never reused or moved


def test_an_enum_has_its_closed_list_and_nothing_else_has_one():
    assert all(bool(r.allowed) == (r.type == "enum") for r in T.RAW_TYPES)


def test_the_spec_rows_for_the_sheet():
    rows = T.spec_rows()
    assert len(rows) == len(T.RAW_TYPES) and all(len(r) == len(T.SPEC_HEADER) for r in rows)
    assert ("Changelog", "CreatedUtc", "datetime", "J2D-DATA-004", "yes", "histories[].created", "") in rows
    assert ("Issues", "SprintState", "enum", "J2D-DATA-011", "no", "fields.<sprint>[].state", "future, active, closed") in rows


# --- the kind of a value ------------------------------------------------------------------------------------------------------

@pytest.mark.parametrize("value, kind", [
    (None, "empty"), ("", "empty"), ("text", "text"), ("1761134538011", "text"), (1761134538011, "number"), (3, "number"), (1.5, "number"),
    (True, "boolean"), (False, "boolean"), (D, "date"), (D.date(), "date"), ([], "other"), ({}, "other"), (" ", "text")])
def test_kind_of(value, kind):
    assert T.kind_of(value) == kind


# --- a value of a type (each type: one fit, one miss per kind that could be taken for it) -------------------------------------------

@pytest.mark.parametrize("type_, value, fits", [
    ("datetime", D, True), ("datetime", dt.datetime(2000, 1, 1), True), ("datetime", dt.datetime(1999, 12, 31, 23, 59), False),
    ("datetime", dt.datetime(2100, 1, 1), False), ("datetime", 1761134538011, False), ("datetime", "2026-11-14T09:30:00", False),
    ("datetime", True, False),
    ("date", dt.datetime(2026, 6, 1), True), ("date", dt.date(2026, 6, 1), True), ("date", D, False), ("date", "2026-06-01", False),
    ("id", "10042", True), ("id", 10042, False), ("id", "ABC-1", False), ("id", "10 042", False), ("id", "١٢", False),
    ("key", "ABC-12", True), ("key", "A1_B-7", True), ("key", "abc-12", False), ("key", "ABC12", False), ("key", 12, False),
    ("integer", 3, True), ("integer", 3.0, True), ("integer", 3.5, False), ("integer", "3", False), ("integer", True, False),
    ("integer", D, False),
    ("number", 3.5, True), ("number", 0, True), ("number", "3.5", False), ("number", True, False), ("number", D, False),
    ("boolean", True, True), ("boolean", False, True), ("boolean", 1, False), ("boolean", "TRUE", False),
    ("text", "To Do", True), ("text", "123", True), ("text", 123, False), ("text", D, False), ("text", True, False),
    ("list", "1.0", True), ("list", "1.0, 2.0", True), ("list", "1.0, , 2.0", False), ("list", "1.0,", False), ("list", 1.5, False),
])
def test_fits(type_, value, fits):
    assert T.fits(type_, value) is fits


def test_an_enum_fits_only_a_value_of_its_list():
    allowed = ("future", "active", "closed")
    assert T.fits("enum", "active", allowed) and not T.fits("enum", "Active", allowed) and not T.fits("enum", "done", allowed)
    assert not T.fits("enum", 1, allowed)


# --- the shape of a value: never the value -------------------------------------------------------------------------------------------

@pytest.mark.parametrize("value, shape", [
    (1761134538011, "Double:9999999999999"), (3.5, "Double:9.9"), ("ABC-12", "String:AAA-99"), ("To Do", "String:Aa Aa"),
    (D, "Date:9999-99-99A99:99:99"), (True, "Boolean"), (None, "Empty"), ("", "Empty"), ([], "Object"),
    ("x" * 50, "String:" + "a" * 40 + "..."), ("café", "String:aaa?")])
def test_shape_of(value, shape):
    assert T.shape_of(value) == shape


def test_a_shape_keeps_no_digit_and_no_letter_of_the_value():
    shape = T.shape_of("PI#26-2026 secret").split(":", 1)[1]
    assert not any(c in shape for c in "secrt26") and shape == "AA#99-9999 aaaaaa"


@pytest.mark.parametrize("value, digits", [(1761134538011, 13), ("1761134538", 10), (12, 2), (2.5, 0), ("12a", 0), (None, 0), (D, 0)])
def test_digits_of(value, digits):
    assert T.digits_of(value) == digits


def test_column_letters():
    assert [T.column_letter(n) for n in (1, 4, 26, 27, 52, 53)] == ["A", "D", "Z", "AA", "AZ", "BA"]


# --- the check ----------------------------------------------------------------------------------------------------------------

HEADER = ("Key", "IssueId", "ChangeId", "CreatedUtc", "CreatedText", "FieldId", "Field", "FieldType", "From", "FromString", "To", "ToString")


def changelog_row(created=D, created_text="2026-11-14T09:30:00.000+0000", key="ABC-1", change="100"):
    return (key, "10001", change, created, created_text, "status", "status", "jira", "1", "To Do", "2", "Done")


def check(rows, header=HEADER, table="Changelog"):
    return T.check({table: (header, rows)})


def test_a_clean_table_has_no_problem_and_counts_its_columns():
    report = check([changelog_row(), changelog_row(change="101")])
    assert report.problems == [] and (report.tables, report.columns, report.checked) == (1, 12, 12)
    assert report.fills["Changelog"][3] == ("CreatedUtc", 2) and report.rows == {"Changelog": 2}


def test_epoch_milliseconds_in_a_date_column_is_the_v070_failure_and_names_its_cause():
    (problem,) = check([changelog_row(created=1761134538011, created_text="1761134538011")]).problems
    assert (problem.name, problem.code, problem.type, problem.found) == ("Changelog.CreatedUtc", "J2D-DATA-004", "datetime", "number")
    assert (problem.cell, problem.rows, problem.filled, problem.kinds) == ("D2", 1, 1, "number x 1")
    assert problem.shape == "Double:9999999999999" and "13 digits: epoch milliseconds" in problem.hint


def test_a_required_date_column_that_is_empty_is_a_problem_found_empty():
    (problem,) = check([changelog_row(created=None), changelog_row(created=None)]).problems
    assert (problem.found, problem.cell, problem.filled, problem.rows, problem.shape) == ("empty", "", 0, 2, "Empty")
    assert "WARN J2D-DATA-003" in problem.hint                                          # where to look next


def test_an_optional_column_that_is_empty_stays_silent():
    row = list(changelog_row())
    row[8:] = [None] * 4                                                                 # From, FromString, To, ToString
    assert check([tuple(row)]).problems == []


def test_the_first_filled_value_is_read_not_the_first_row():
    header = ("Key", "Id", "Project", "Type", "WorkType", "CreatedUtc", "Status", "Parent", "StatusCategory", "ResolvedUtc", "FixVersions",
              "StoryPoints", "PI", "BV", "TeamValue", "SprintState", "SprintStart", "SprintEnd", "SprintCount", "TargetStart", "TargetEnd",
              "KeyChanges")

    def issue(parent=None, resolved=None, points=None):
        return ("ABC-1", "1", "ABC", "Story", None, D, "Done", parent, "done", resolved, None, points, None, None, None, "", None, None, 0,
                None, None, 0)
    report = T.check({"Issues": (header, [issue(), issue(), issue(parent="ABC-9", resolved=D, points=5.0)])})
    assert report.problems == []
    assert dict(report.fills["Issues"])["ResolvedUtc"] == 1 and dict(report.fills["Issues"])["StoryPoints"] == 1


def test_a_text_in_a_number_column_is_found_and_the_hint_points_at_the_field_id():
    header = ("Key", "Id", "Project", "Type", "WorkType", "CreatedUtc", "Status", "Parent", "StatusCategory", "ResolvedUtc", "FixVersions",
              "StoryPoints", "PI", "BV", "TeamValue", "SprintState", "SprintStart", "SprintEnd", "SprintCount", "TargetStart", "TargetEnd",
              "KeyChanges")
    row = ("ABC-1", "1", "ABC", "Story", None, D, "Done", None, "done", None, None, "5 points", None, None, None, "", None, None, 0, None, None, 0)
    (problem,) = T.check({"Issues": (header, [row])}).problems
    assert (problem.code, problem.found, problem.cell) == ("J2D-DATA-009", "text", "L2")
    assert "custom field is of another type" in problem.hint and problem.shape == "String:9 aaaaaa"


def test_a_missing_column_is_a_problem_of_its_type():
    problems = check([changelog_row()[:3]], header=HEADER[:3]).problems
    assert [p.name for p in problems] == [f"Changelog.{c}" for c in HEADER[3:]] and all(p.found == "missing" for p in problems)
    absent = next(p for p in problems if p.column == "CreatedUtc")
    assert (absent.found, absent.code) == ("missing", "J2D-DATA-004") and "not in the sheet" in absent.hint


def test_a_table_without_a_row_or_not_given_is_not_read():
    report = T.check({"Changelog": (HEADER, []), "Other": (("A",), [(1,)])})
    assert report == T.TypeReport([], 0, 0, 0, {}, {})


def test_the_kinds_of_the_first_twenty_filled_values_are_counted():
    rows = [changelog_row(created=1761134538011 + n) for n in range(25)]
    rows[3] = changelog_row(created="x")
    (problem,) = check(rows).problems
    assert problem.kinds == "number x 19, text x 1" and problem.filled == 25


def test_every_type_gives_its_own_code_on_a_wrong_value():
    wrong = {"datetime": ("Changelog", "CreatedUtc", "x"), "date": ("Issues", "TargetStart", "x"), "id": ("Changelog", "IssueId", 5),
             "key": ("Changelog", "Key", "abc"), "integer": ("Issues", "KeyChanges", 1.5), "number": ("Issues", "BV", "x"),
             "boolean": ("StatusIntervals", "IsOpen", "yes"), "enum": ("Issues", "SprintState", "paused"),
             "list": ("Issues", "FixVersions", "1.0,,2.0"), "text": ("Changelog", "FieldId", 7)}
    assert set(wrong) == set(T.TYPE_CODES)
    for type_, (table, column, value) in wrong.items():
        rule = next(r for r in T.RAW_TYPES if r.table == table and r.column == column)
        assert rule.type == type_ and not T.fits(type_, value, rule.allowed), (type_, column)
        assert T.hint(rule, T.kind_of(value), value)                                    # a hint for each, never empty


# --- the VBA side says the same words (vba/modHealth.bas: HintFor) -------------------------------------------------------------------------

HEALTH = (Path(__file__).resolve().parents[2] / "vba" / "modHealth.bas").read_text(encoding="ascii")


def normal(text: str) -> str:
    return re.sub(r"^\d+ digits", "N digits", text)


def vba_hint_texts() -> set[str]:
    """Each `HintFor = <pieces>` of modHealth: the literals joined, `digits` as 0 and `allowed` as A."""
    texts = set()
    for expression in re.findall(r"HintFor = (.*?)(?::\s*Exit Function)?\s*$", HEALTH, re.MULTILINE):
        parts = re.findall(r'"((?:[^"]|"")*)"|\b(digits|allowed)\b', expression)
        texts.add(normal("".join(lit.replace('""', '"') if not name else {"digits": "0", "allowed": "A"}[name] for lit, name in parts)))
    return texts


def python_hint_texts() -> set[str]:
    texts = set()
    for rule_type in T.TYPE_CODES:
        for required in (True, False):
            rule = T.Rule("T", "C", rule_type, required, "src", ("A",) if rule_type == "enum" else ())
            for found, first in (("missing", None), ("empty", None), ("number", 1761134538011), ("number", 3.5), ("text", "1761134538"),
                                 ("text", "x"), ("date", D), ("boolean", True)):
                texts.add(normal(T.hint(rule, found, first)))
    return texts


def test_the_hints_of_the_vba_are_the_hints_of_python():
    """The two sides give the same sentences: one changed on a single side fails here."""
    vba, python = vba_hint_texts(), python_hint_texts()
    assert len(vba) == 19
    assert vba == python, (sorted(vba - python), sorted(python - vba))


# --- the lines of the log: names, counts, shapes, no value ------------------------------------------------------------------------------

def test_a_problem_makes_three_lines_that_hold_everything_to_debug_and_no_value():
    secret = 1761134538011
    (problem,) = check([changelog_row(created=secret, created_text="x")]).problems
    (warn, counts, hint) = problem.lines()
    assert warn == ("WARN", "J2D-DATA-004 | Changelog.CreatedUtc | expected datetime | found number | source histories[].created | first filled D2")
    assert counts == ("INFO", "J2D-DATA-004 | Changelog.CreatedUtc | rows 1, filled 1, empty 0 | kinds in first 20: number x 1 | "
                              "value shape Double:9999999999999")
    assert hint[0] == "INFO" and hint[1].startswith("J2D-DATA-004 | Changelog.CreatedUtc | hint: 13 digits: epoch milliseconds")
    assert all(str(secret) not in text and "\n" not in text for _level, text in problem.lines())


def test_every_line_of_a_problem_fits_the_300_characters_of_the_vba_log_in_the_worst_case():
    worst = T.TypeProblem("StatusIntervals", "StatusCategoryOfTheTicket", "enum", "J2D-DATA-011", "boolean", "fields.status.statusCategory.key",
                          "AB123456", 1_048_575, 1_048_574, "boolean x 4, date x 4, error x 4, number x 4, other x 2, text x 2", "String:" + "a" * 40 + "...",
                          T.hint(T.Rule("T", "C", "enum", False, "s", ("future", "active", "closed")), "text"))
    assert all(len(text) < 300 for _level, text in worst.lines())


def test_the_summary_and_fill_lines_and_the_status_field():
    report = check([changelog_row(created=None), changelog_row(created=None)])
    assert T.summary_line(report) == "column types | tables 1 | columns 12 | checked 11 | problems 1"
    assert T.fill_lines(report)[0].startswith("types Changelog | rows 2 | filled: Key 2, IssueId 2, ChangeId 2, CreatedUtc 0, CreatedText 2")
    assert T.status_field(report) == " | Column types: 1 column(s), first Changelog.CreatedUtc expected datetime found empty"
    assert T.status_field(check([changelog_row()])) == ""


def test_the_fill_counts_of_a_wide_table_go_on_a_next_line_so_that_no_line_passes_the_log_limit():
    columns = tuple(r.column for r in T.RAW_TYPES if r.table == "Issues")
    report = T.TypeReport([], 1, len(columns), len(columns), {"Issues": [(c, 12345) for c in columns]}, {"Issues": 99999})
    lines = T.fill_lines(report)
    assert len(lines) == 2 and all(len(line) <= 300 for line in lines)
    assert lines[0].startswith("types Issues | rows 99999 | filled: Key 12345, Id 12345") and lines[1].startswith("types Issues | filled (continued): ")
    joined = ", ".join(line.split(": ", 1)[1] for line in lines)
    assert joined == ", ".join(f"{c} 12345" for c in columns)                           # nothing lost, nothing twice


def test_the_wide_table_of_the_vba_test_splits_at_the_same_place():
    """T_TypeFillLines: 22 columns of 12 characters, 2 rows: 15 on the first line, 7 on the next."""
    columns = [(f"WideColumn{n:02d}", 2) for n in range(1, 23)]
    lines = T.fill_lines(T.TypeReport([], 1, 22, 22, {"Wide": columns}, {"Wide": 2}))
    assert lines == ["types Wide | rows 2 | filled: " + ", ".join(f"WideColumn{n:02d} 2" for n in range(1, 16)),
                     "types Wide | filled (continued): " + ", ".join(f"WideColumn{n:02d} 2" for n in range(16, 23))]


def test_the_log_lines_come_per_table_problems_first_then_the_fill_lines_then_the_summary():
    """The order modHealth.TypeProblems writes them (T_TypeProblems checks it on the VBA side): the tables in the order of the spec."""
    report = T.check({"Changelog": (HEADER, [changelog_row(created=None)]), "IssueLinks": (("Blocker", "Blocked"), [("ABC-1", "x")])})
    tags = []
    for level, text in T.log_lines(report):
        if level == "WARN":
            tags.append("W:" + text.split(" | ")[1])
        elif text.startswith("types "):
            tags.append("F:" + text.split(" | ")[0][len("types "):])
        elif text.startswith("column types"):
            tags.append("S")
    assert tags == ["W:IssueLinks.Blocked", "F:IssueLinks", "W:Changelog.CreatedUtc", "F:Changelog", "S"]
    assert T.log_lines(report)[-1] == ("INFO", "column types | tables 2 | columns 14 | checked 13 | problems 2")
    assert [level for level, _text in T.log_lines(report)[:3]] == ["WARN", "INFO", "INFO"]                # three lines per problem


@pytest.mark.parametrize("value, shape", [(0.5, "Double:9.9"), (0.1 + 0.2, "Double:9.9"), (1e15, "Double:" + "9" * 16), (-3, "Double:-9"), (-0.5, "Double:-9.9")])
def test_the_shape_of_a_number_is_the_same_in_every_locale_and_on_both_sides(value, shape):
    assert T.shape_of(value) == shape


def test_a_status_without_a_category_is_a_valid_value():
    rule = next(r for r in T.RAW_TYPES if r.column == "StatusCategory")
    assert all(T.fits("enum", value, rule.allowed) for value in ("new", "indeterminate", "done", "undefined"))
    assert not T.fits("enum", "other", rule.allowed)


def test_a_table_with_no_column_to_count_makes_no_fill_line():
    assert T.fill_lines(T.TypeReport([], 1, 0, 0, {"Issues": []}, {"Issues": 3})) == []


def test_the_status_field_shows_five_names_and_says_there_are_more():
    header = tuple(r.column for r in T.RAW_TYPES if r.table == "Changelog")
    report = check([(1,) * len(header)], header=header)
    assert len(report.problems) > 5
    field = T.status_field(report)
    assert field.count(" expected ") == 5 and field.endswith("; ...") and f"{len(report.problems)} column(s)" in field
