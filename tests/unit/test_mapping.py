"""Jira's JSON as the raw sheets and the flow issues of the Python path (pure).

The oracle: the fake Jira's answers (mockjira.api, +0200 offset), mapped, must give the same flow issues as
flow_adapter.to_flow_issues on the same data, to the second (the VBA drops the milliseconds too).
"""
import dataclasses
import datetime as dt

import pytest

from jira2dashboard.flow.formulas import CHANGELOG_RAW, INTERVALS_RAW, ISSUES
from jira2dashboard.flow.model import ART, Change, FlowIssue, Rules
from jira2dashboard.mock_workbook import FIELD_IDS
from jira2dashboard.mockjira import api
from jira2dashboard.mockjira.anomalies import with_anomalies
from jira2dashboard.mockjira.flow_adapter import to_flow_issues
from jira2dashboard.mockjira.generator import generate
from jira2dashboard.params.model import load_all, project_keys
from jira2dashboard.standalone import jira_api as J
from jira2dashboard.standalone import mapping as M

PROJECTS = [row for s in load_all() if s.name == "Projects" for row in s.rows]


def answers(dataset, tz=120):
    """The search pages and the changelog pages the extraction would read from the fake Jira."""
    fields = {name: getattr(dataset.fields, attr) for name, attr in FIELD_IDS.items()}
    ids = M.FieldIds(dataset.fields.work_type, fields)
    query = {"jql": J.projects_jql(PROJECTS), "fields": J.search_fields(ids.work_type, fields), "maxResults": 100}
    issues, token = [], None
    while True:
        page = api.search(dataset, query | ({"nextPageToken": token} if token else {}), tz=tz, base_url="x",
                          page_clamp=100)
        issues += page["issues"]
        token = page["nextPageToken"]
        if page["isLast"]:
            break
    pages = []
    for chunk in J.chunks([i["key"] for i in issues]):
        token = None
        while True:
            body = {"issueIdsOrKeys": chunk, "maxResults": 1000} | ({"nextPageToken": token} if token else {})
            page = api.bulk_changelog(dataset, body, tz=tz, bulk_clamp=300)
            pages.append(page)
            token = page["nextPageToken"]
            if not token:
                break
    return issues, pages, ids


def to_second(issue: FlowIssue) -> FlowIssue:
    """The reference without milliseconds, its changes in date order."""
    def floor(value):
        return value.replace(microsecond=0) if isinstance(value, dt.datetime) else value
    changes = tuple(sorted((Change(c.id, floor(c.at), c.from_status, c.to_status) for c in issue.changes),
                           key=lambda c: (c.at, c.id)))
    return dataclasses.replace(issue, **{f.name: floor(getattr(issue, f.name)) for f in dataclasses.fields(issue)
                                         if isinstance(getattr(issue, f.name), dt.datetime)}, changes=changes)


def test_the_flow_issues_of_the_json_are_those_of_the_fake_jira():
    dataset, cases = with_anomalies(generate(seed=11, issues=150, projects=project_keys()), project_keys())
    issues, pages, ids = answers(dataset)
    extract = M.extracted(issues, pages, ids, PROJECTS, "Blocks")
    got = {i.key: dataclasses.replace(i, changes=tuple(sorted(i.changes, key=lambda c: (c.at, c.id))))
           for i in extract.flow}
    expected = {i.key: to_second(i) for i in to_flow_issues(dataset)}
    assert set(got) == set(expected)
    wrong = [(key, f.name) for key in expected for f in dataclasses.fields(FlowIssue)
             if getattr(got[key], f.name) != getattr(expected[key], f.name)]
    assert wrong == []
    assert set(cases) <= set(got) and any(i.blocks for i in got.values()) and any(i.key_changes for i in got.values())
    assert extract.changelog_items == sum(len(i.changes) + len(i.key_changes) for i in dataset.issues)


def test_the_rows_have_the_columns_of_the_sheets_and_the_program_comes_first():
    dataset = generate(seed=3, issues=40, projects=project_keys())
    issues, pages, ids = answers(dataset)
    extract = M.extracted(issues, pages, ids, PROJECTS, "Blocks")
    assert all(len(row) == len(ISSUES) for row in extract.rows)
    program = M.program_project(PROJECTS)
    projects = [row[2] for row in extract.rows]
    assert projects == sorted(projects, key=lambda p: p != program) and projects[0] == program
    assert all(isinstance(row[1], str) for row in extract.rows)                       # Id is text, as in the VBA
    assert [i.team for i in extract.flow if i.workflow == "Feature"] == [ART] * projects.count(program)


@pytest.mark.parametrize("text, moment", [
    ("2026-10-01T09:12:33.123+0200", dt.datetime(2026, 10, 1, 7, 12, 33)),
    ("2026-10-01T09:12:33.123+02:00", dt.datetime(2026, 10, 1, 7, 12, 33)),
    ("2026-10-01T09:12:33.000-0530", dt.datetime(2026, 10, 1, 14, 42, 33)),
    ("2026-10-01T09:12:33Z", dt.datetime(2026, 10, 1, 9, 12, 33)),
    ("2026-10-01", dt.datetime(2026, 10, 1)),
    ("2026-13-01", None), ("01/10/2026", None), ("2026-10", None), ("", None), (None, None), ({"a": 1}, None),
    ("2026-10-01T09:12:33+02", None),
])
def test_jira_dates_become_utc_like_in_the_vba(text, moment):
    assert M.parse_jira_date(text) == moment


def test_text_number_and_work_type_like_in_the_vba():
    assert [M.text_of(v) for v in (None, [1], {"name": "Alpha", "value": "x"}, {"value": "V"}, {"id": 1}, True, 5.0,
                                   2.5, 7, "PI#01-2026")] == [
        None, None, "Alpha", "V", None, "true", "5", "2.5", "7", "PI#01-2026"]
    assert [M.number_of(v) for v in (None, True, {"v": 1}, [1], 3, 2.5, " 8 ", "1.5.2", "+", "8 pts", "")] == [
        None, None, None, None, 3.0, 2.5, 8.0, 1.5, 0.0, None, None]
    fields = {"cf": {"value": "Enabler"}, "plain": "Defect", "list": ["x"], "none": None}
    assert [M.work_type_of(fields, f) for f in ("cf", "plain", "list", "none", "missing", "")] == [
        "Enabler", "Defect", None, None, None, None]


def test_the_sprint_that_counts():
    def sprint(state, day):
        return {"state": state, "startDate": f"2026-10-{day:02d}T00:00:00.000+0000",
                "endDate": f"2026-10-{day + 13:02d}T00:00:00.000+0000"}
    closed1, closed2, future, active = sprint("closed", 1), sprint("closed", 5), sprint("future", 9), sprint("active", 3)
    assert M.sprint_info(None) == ("", None, None, 0)
    assert M.sprint_info([closed1, closed2]) == ("closed", dt.datetime(2026, 10, 5), dt.datetime(2026, 10, 18), 2)
    assert M.sprint_info([future, closed1, active, closed2])[0:2] == ("active", dt.datetime(2026, 10, 3))
    assert M.sprint_info([closed1, future])[0] == "future" and M.sprint_info(active)[3] == 1
    assert M.sprint_info([{"state": "unknown"}]) == ("", None, None, 1)


def test_a_link_counts_once_from_the_blocker_and_only_of_the_type_of_config():
    issues = [{"key": "A-1", "fields": {"issuelinks": [
        {"type": {"name": "Blocks"}, "outwardIssue": {"key": "B-1"}},
        {"type": {"name": "Blocks"}, "inwardIssue": {"key": "C-1"}},
        {"type": {"name": "Relates"}, "outwardIssue": {"key": "D-1"}}]}}]
    assert M.block_links(issues, "Blocks") == [("A-1", "B-1")]
    assert M.block_links(issues, "") == [] and M.block_links(issues, "blocks") == []


def issue(changes, status="Terminé", created=dt.datetime(2026, 10, 1)):
    return FlowIssue("TA-1", "Alpha", "Story", "Feature", "Team", created, status,
                     tuple(Change(str(n), at, a, b) for n, (at, a, b) in enumerate(changes, start=1)))


def test_the_status_intervals_name_what_is_inconsistent():
    day = dt.datetime(2026, 10, 2)
    rows = M.interval_rows(issue([(day, "Nouveau", "En analyse"), (day + dt.timedelta(1), "Prêt", "Terminé"),
                                  (dt.datetime(2026, 9, 30), "Terminé", "En test")], status="Terminé"),
                           dt.datetime(2026, 11, 3), Rules.from_params())
    assert all(len(row) == len(INTERVALS_RAW) for row in rows)
    assert [(r[2], r[3], r[10]) for r in rows] == [                   # in date order: the change of 30 Sept first
        (1, "Terminé", "BEFORE_CREATED"), (2, "En test", "FROM_MISMATCH(prev=En test)"),
        (3, "En analyse", "FROM_MISMATCH(prev=En analyse)"), (4, "Terminé", None)]
    assert rows[-1][8] is True and rows[-1][6] is None and rows[0][1] == "Team"
    assert M.interval_checks(issue([], status="Nouveau")) == [""]
    assert M.interval_checks(issue([(day, "Nouveau", "Prêt")], status="Terminé")) == ["", "CURRENT_MISMATCH(now=Terminé)"]


def test_status_and_key_changes_come_from_their_items():
    pages = [{"issueChangeLogs": [{"issueId": "10", "changeHistories": [
        {"id": "h1", "created": "2026-10-02T12:00:00.000+0200", "items": [
            {"fieldId": "status", "fromString": "Nouveau", "toString": "Prêt"},
            {"field": "assignee", "fieldId": "assignee"}]},
        {"id": "h2", "created": "2026-10-03T12:00:00.000+0200",
         "items": [{"field": "Key", "fieldId": "issuekey", "fromString": "OLD-1", "toString": "TA-1"}]}]}]}]
    items = M.changelog_items(pages)
    assert len(items) == 3
    assert M.status_changes(items, {"10": "TA-1"}) == {"TA-1": [Change("h1", dt.datetime(2026, 10, 2, 10), "Nouveau",
                                                                        "Prêt")]}
    assert M.key_change_counts(items, {"10": "TA-1"}) == {"TA-1": 1}


# --- status changes whose date cannot be read: left out, counted, and what the value looks like (J2D-DATA-003) ---------------

@pytest.mark.parametrize("value, shape", [
    ("2026-10-01T09:12:33.123+0200", "String:9999-99-99A99:99:99.999+9999"),      # the same texts as modFlow.DateShape (T_DateShape)
    ("Hello World", "String:Aaaaa Aaaaa"),
    ("1696000000000", "String:9999999999999"),
    ("eé", "String:a?"),
    ("7" * 50, "String:" + "9" * 40 + "..."),
    (None, "Null"),
    ({"a": 1}, "Object"),
])
def test_the_shape_of_a_date_that_cannot_be_read_holds_none_of_its_content(value, shape):
    assert M.date_shape(value) == shape


def test_status_changes_without_a_readable_date_are_counted_with_their_shapes():
    items = [("10", {"id": "h1", "created": "2026-10-02T12:00:00.000+0200"}, {"fieldId": "status"}),
             ("10", {"id": "h2", "created": "1696000000000"}, {"fieldId": "status"}),
             ("10", {"id": "h3", "created": "1696000000001"}, {"fieldId": "status"}),
             ("10", {"id": "h4"}, {"fieldId": "status"}),
             ("10", {"id": "h5", "created": "not a date"}, {"fieldId": "assignee"})]          # not a status: not counted
    assert M.unreadable_status_dates(items) == (3, "String:9999999999999 x 2; Null x 1")
    assert M.unreadable_status_dates(items[:1]) == (0, "")

def test_the_extract_tells_how_many_status_changes_were_left_out():
    pages = [{"issueChangeLogs": [{"issueId": "10", "changeHistories": [
        {"id": "h1", "created": "2026-10-02T12:00:00.000+0200", "items": [{"fieldId": "status", "fromString": "A", "toString": "B"}]},
        {"id": "h2", "created": "1696000000000", "items": [{"fieldId": "status", "fromString": "B", "toString": "C"}]}]}]}]
    extract = M.extracted([], pages, M.FieldIds("", {}), [], "Blocks")                # no issue: only the changelog matters here
    assert (extract.unreadable_dates, extract.unreadable_shapes) == (1, "String:9999999999999 x 1")
    assert extract.changelog_items == 2


# --- the sheet Changelog: every item, no author (modFlow.ChangelogRows) -------------------------------------------------------------

def test_the_changelog_keeps_every_item_without_author_cuts_the_texts_and_keeps_no_person():
    items = [("10", {"id": "h1", "created": "2026-10-02T12:00:00.000+0200"},
              {"fieldId": "status", "field": "status", "fieldtype": "jira", "from": "1", "fromString": "Nouveau", "to": "3", "toString": "Pret"}),
             ("10", {"id": "h2", "created": "1696000000000"},
              {"fieldId": "description", "field": "description", "fieldtype": "jira", "fromString": "x" * 300, "toString": "ok"}),
             ("11", {"id": "h3", "created": "2026-10-03T10:00:00.000Z", "author": {"displayName": "Ann"}},
              {"fieldId": "assignee", "field": "assignee", "fieldtype": "jira", "from": "acc-1", "fromString": "Mary Martin",
               "to": "acc-2", "toString": "Paul Durand"})]
    rows = M.changelog_rows(items, {"10": "TA-1"})
    assert rows[0] == ("TA-1", "10", "h1", dt.datetime(2026, 10, 2, 10), "2026-10-02T12:00:00.000+0200", "status", "status", "jira",
                       "1", "Nouveau", "3", "Pret")
    assert rows[1][3] is None and rows[1][4] == "1696000000000" and len(rows[1][9]) == 255      # not read, shown; cut
    assert rows[2][0] == "11" and rows[2][8:] == (None, None, None, None)                       # an unknown id stays; a person keeps no value
    assert all(len(row) == len(CHANGELOG_RAW) for row in rows)
    assert not any(name in str(rows) for name in ("Ann", "Martin", "Durand"))


def test_the_changelog_columns_are_those_of_the_vba_header():
    import re
    from pathlib import Path
    text = (Path(__file__).resolve().parents[2] / "vba" / "modFlow.bas").read_text(encoding="ascii")
    body = text.split("Public Function ChangelogHeader() As Variant")[1].split("End Function")[0]
    assert tuple(re.findall(r'"(\w+)"', body)) == CHANGELOG_RAW


def test_the_changelog_hides_every_field_jira_types_as_a_user_and_the_mentions_in_the_texts():
    items = [("10", {"id": "h1", "created": "2026-10-02T12:00:00.000+0200"},
              {"fieldId": "customfield_10050", "field": "Product owner", "fieldtype": "custom", "from": "acc-9", "fromString": "Mary Martin",
               "to": "acc-8", "toString": "Paul Durand"}),
             ("10", {"id": "h2", "created": "2026-10-03T12:00:00.000+0200"},
              {"fieldId": "description", "field": "description", "fieldtype": "jira",
               "fromString": "Ask [~accountid:5b10ac8d82e05b22cc7d4ef5] and [~jdoe] now", "toString": "ok"})]
    rows = M.changelog_rows(items, {"10": "TA-1"}, frozenset({"customfield_10050"}))
    assert rows[0][8:] == (None, None, None, None)                                      # a custom people picker: no value
    assert rows[1][9] == "Ask [user] and [user] now"
    assert not any(name in str(rows) for name in ("Martin", "Durand", "5b10", "jdoe"))
    assert M.changelog_rows(items[:1], {"10": "TA-1"})[0][9] == "Mary Martin"          # without the list, only the fixed fields are hidden
