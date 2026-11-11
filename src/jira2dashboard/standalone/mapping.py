"""Jira's JSON as the raw sheets and as flow issues, pure: what the VBA modules modExtract and modFlow do, in Python.

Issues (the 22 columns of formulas.ISSUES), StatusIntervals (the 11 raw columns, with the Check of each visit),
IssueLinks (blocker, blocked) and the FlowIssue the flow reference computes from. The conversions are those of the
VBA: a date with an offset becomes UTC, a select field gives its value, a team its name, the sprint that counts is
the active one, else a future one, else the last closed one.
"""
import datetime as dt
import re
from dataclasses import dataclass

from jira2dashboard.flow.intervals import build_intervals
from jira2dashboard.flow.model import ART, Change, FlowIssue, Rules

_NUMBER_TEXT = re.compile(r"[0-9.+-]+")
_LEADING_NUMBER = re.compile(r"[+-]?(\d+\.?\d*|\.\d+)")
SPRINT_RANK = {"active": 3, "future": 2, "closed": 1}


def parse_jira_date(value) -> dt.datetime | None:
    """2026-10-01T09:12:33.123+0200, ...+02:00, ...Z -> naive UTC (milliseconds dropped); 2026-10-01 -> that day.
    Empty or invalid -> None, never an error (modText.ParseJiraDate)."""
    if value is None or isinstance(value, (dict, list, bool)):
        return None
    text = str(value)
    if len(text) < 10 or text[4] != "-" or text[7] != "-":
        return None
    try:
        day = dt.datetime(int(text[:4]), int(text[5:7]), int(text[8:10]))
        if len(text) == 10:
            return day
        moment = day + dt.timedelta(hours=int(text[11:13]), minutes=int(text[14:16]), seconds=int(text[17:19]))
        tail = text[19:]
        while tail and tail[0] not in "+-Zz":
            tail = tail[1:]                                          # skip ".123"
        offset = 0
        if tail and tail[0] not in "Zz":
            digits = tail[1:].replace(":", "")
            offset = (-1 if tail[0] == "-" else 1) * (int(digits[:2]) * 60 + int(digits[2:4]))
        return moment - dt.timedelta(minutes=offset)
    except (ValueError, IndexError):
        return None


def text_of(value) -> str | None:
    """A text field, or an object (Team {"name"}, a select {"value"}): the text; a list or nothing: None."""
    if value is None or isinstance(value, list):
        return None
    if isinstance(value, dict):
        found = value.get("name")
        return found if found is not None else value.get("value")
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else repr(value)
    return str(value)


def number_of(value) -> float | None:
    """A number field. Text only as plain digits with a dot (never the decimal separator of the PC); not a Boolean."""
    if value is None or isinstance(value, (bool, dict, list)):
        return None
    if isinstance(value, str):
        cleaned = value.strip()
        if not cleaned or not _NUMBER_TEXT.fullmatch(cleaned):
            return None
        found = _LEADING_NUMBER.match(cleaned)                        # Val("1.5.2") is 1.5, Val("+") is 0
        return float(found.group(0)) if found else 0.0
    return float(value)


def work_type_of(fields: dict, field_id: str):
    """The work type (TypeCategorie): a select field comes as {"value": ...}; None when not set."""
    if not field_id:
        return None
    value = fields.get(field_id)
    if isinstance(value, dict):
        return value.get("value")
    return None if isinstance(value, list) else value


def date_of(value) -> dt.datetime | None:
    return None if value is None or isinstance(value, (dict, list)) else parse_jira_date(str(value))


def versions_of(fields: dict) -> tuple[str, ...]:
    return tuple(str(v.get("name") or "") for v in fields.get("fixVersions") or [] if isinstance(v, dict))


def sprint_info(value) -> tuple[str, dt.datetime | None, dt.datetime | None, int]:
    """(state, start, end, count): the active sprint, else a future one, else the last closed one (the last of the
    best state in the list); count = every sprint of the field. No sprint: ("", None, None, 0)."""
    sprints = value if isinstance(value, list) else [value] if isinstance(value, dict) else []
    best, best_rank = ("", None, None), 0
    for sprint in sprints:
        state = str((sprint or {}).get("state") or "")
        rank = SPRINT_RANK.get(state, 0)
        if rank and rank >= best_rank:
            best_rank = rank
            best = (state, parse_jira_date(sprint.get("startDate")), parse_jira_date(sprint.get("endDate")))
    return best + (len(sprints),)


class FieldIds:
    """The custom fields of Config: WorkTypeField and the ...Field of the data checks ("" = not configured)."""

    def __init__(self, work_type: str, fields: dict[str, str]):
        self.work_type = work_type
        self.story_points = fields.get("StoryPointsField", "")
        self.business_value = fields.get("BusinessValueField", "")
        self.pi = fields.get("PIField", "")
        self.team = fields.get("TeamField", "")
        self.sprint = fields.get("SprintField", "")
        self.target_start = fields.get("TargetStartField", "")
        self.target_end = fields.get("TargetEndField", "")


def _field(fields: dict, field_id: str):
    return fields.get(field_id) if field_id else None


def issue_row(issue: dict, ids: FieldIds, key_changes: dict[str, int]) -> tuple:
    """One row of the sheet Issues (modExtract.IssueRowOf), in the order of formulas.ISSUES."""
    fields = issue.get("fields") or {}
    key = issue.get("key")
    status = fields.get("status") or {}
    sprint = sprint_info(_field(fields, ids.sprint))
    versions = versions_of(fields)
    return (key, str(issue.get("id") or ""), (fields.get("project") or {}).get("key"),
            (fields.get("issuetype") or {}).get("name"), work_type_of(fields, ids.work_type),
            parse_jira_date(fields.get("created")), status.get("name"), text_of((fields.get("parent") or {}).get("key")),
            text_of((status.get("statusCategory") or {}).get("key")), parse_jira_date(fields.get("resolutiondate")),
            ", ".join(versions) or None, number_of(_field(fields, ids.story_points)), text_of(_field(fields, ids.pi)),
            number_of(_field(fields, ids.business_value)), text_of(_field(fields, ids.team)),
            sprint[0] or None, sprint[1], sprint[2], sprint[3],
            date_of(_field(fields, ids.target_start)), date_of(_field(fields, ids.target_end)),
            key_changes.get(key, 0))


def block_links(issues: list[dict], link_type: str) -> list[tuple[str, str]]:
    """(blocker, blocked) for each link of type `link_type` that starts at the issue (outward side only: the other end
    shows the same link as inward, so each link counts once). No link type: no row (modExtract.AddBlockLinks)."""
    if not link_type:
        return []
    out = []
    for issue in issues:
        for link in (issue.get("fields") or {}).get("issuelinks") or []:
            other = (link.get("outwardIssue") or {}).get("key")
            if str((link.get("type") or {}).get("name") or "") == link_type and other:
                out.append((issue.get("key"), other))
    return out


def changelog_items(pages: list[dict]) -> list[tuple[str, dict, dict]]:
    """(issue id, history, item) of every item of the changelog pages, in the order Jira sent them."""
    return [(str(log.get("issueId")), history, item)
            for page in pages for log in page.get("issueChangeLogs") or []
            for history in log.get("changeHistories") or [] for item in history.get("items") or []]


def status_changes(items, id_to_key: dict[str, str]) -> dict[str, list[Change]]:
    """{key: status changes} from the items whose fieldId is status (modFlow.StatusChangesFromChangelog)."""
    out: dict[str, list[Change]] = {}
    for issue_id, history, item in items:
        key = id_to_key.get(issue_id, issue_id)
        if item.get("fieldId") == "status":
            at = parse_jira_date(history.get("created"))
            if at is not None:
                out.setdefault(key, []).append(Change(str(history.get("id")), at, item.get("fromString") or "",
                                                      item.get("toString") or ""))
    return out


def key_change_counts(items, id_to_key: dict[str, str]) -> dict[str, int]:
    """{key: moves between projects}: the items of the field Key (modFlow.KeyChangeCounts)."""
    out: dict[str, int] = {}
    for issue_id, _history, item in items:
        if item.get("field") == "Key" or item.get("fieldId") == "issuekey":
            key = id_to_key.get(issue_id, issue_id)
            out[key] = out.get(key, 0) + 1
    return out


def program_project(projects: list[tuple[str, str, str]]) -> str:
    return next((row[0] for row in projects if (row[1] or "").strip().lower() == "program"), "")


def flow_issue(row: tuple, changes: list[Change], projects: list[tuple[str, str, str]], blocks: tuple[str, ...],
               fix_versions: tuple[str, ...]) -> FlowIssue:
    """The issue as the flow reference reads it: Program issues follow the Feature workflow and report under ART,
    the others follow the Team workflow under the team of their project (sheet Projects)."""
    (key, _id, project, type_, work_type, created, status, parent, category, resolved, _versions, points, pi, value,
     team_value, sprint_state, sprint_start, sprint_end, sprint_count, target_start, target_end, key_changes) = row
    program = project == program_project(projects)
    team = next((row_[2] or "" for row_ in projects if row_[0] == project), "")
    return FlowIssue(key=key, team=ART if program else team, type=type_, work_type=work_type,
                     workflow="Feature" if program else "Team", created=created, status=status,
                     changes=tuple(changes), parent=parent, resolved=resolved, status_category=category or "",
                     fix_versions=fix_versions,
                     story_points=points, pi=pi, business_value=value, team_value=team_value,
                     sprint_state=sprint_state or "", sprint_start=sprint_start, sprint_end=sprint_end,
                     sprint_count=sprint_count, target_start=target_start, target_end=target_end, blocks=blocks,
                     key_changes=key_changes)


def interval_checks(issue: FlowIssue) -> list[str]:
    """The Check of each visit, in the order of build_intervals: FROM_MISMATCH(prev=...) when a change does not leave
    the status the issue was in, BEFORE_CREATED when it is dated before the visit started, CURRENT_MISMATCH(now=...)
    on the open visit when the history does not end in the current status (modFlow.AddIssueIntervals)."""
    changes = sorted(issue.changes, key=lambda c: (c.at, c.id))
    running = changes[0].from_status if changes else issue.status
    entered = issue.created
    out = []
    for change in changes:
        marks = []
        if change.from_status != running:
            marks.append(f"FROM_MISMATCH(prev={running})")
        if change.at < entered:
            marks.append("BEFORE_CREATED")
        out.append(";".join(marks))
        running, entered = change.to_status, change.at
    out.append(f"CURRENT_MISMATCH(now={issue.status})" if running != issue.status else "")
    return out


def interval_rows(issue: FlowIssue, as_of: dt.datetime, rules: Rules) -> list[tuple]:
    """The rows of the sheet StatusIntervals (formulas.INTERVALS_RAW) of one issue."""
    return [(i.key, issue.workflow, i.seq, i.status, i.category, i.entered, i.left, i.duration_days, i.is_open,
             i.change_id, check or None)
            for i, check in zip(build_intervals(issue, as_of, rules), interval_checks(issue))]


@dataclass
class Extract:
    """What the extraction read, as rows and as flow issues."""
    rows: list[tuple]                       # sheet Issues
    flow: list[FlowIssue]
    links: list[tuple[str, str]]            # sheet IssueLinks
    changelog_items: int


def extracted(issues: list[dict], changelog_pages: list[dict], ids: FieldIds, projects: list[tuple[str, str, str]],
              link_type: str) -> Extract:
    """The issues of the search pages with their changelogs. The Program issues come first, as the VBA writes them."""
    id_to_key = {str(i.get("id")): i.get("key") for i in issues}
    items = changelog_items(changelog_pages)
    changes, moved = status_changes(items, id_to_key), key_change_counts(items, id_to_key)
    links = block_links(issues, link_type)
    program = program_project(projects)
    ordered = sorted(issues, key=lambda i: ((i.get("fields") or {}).get("project") or {}).get("key") != program)
    blocked_by: dict[str, list[str]] = {}
    for blocker, blocked in links:
        blocked_by.setdefault(blocker, []).append(blocked)
    rows, flow = [], []
    for issue in ordered:
        row = issue_row(issue, ids, moved)
        rows.append(row)
        blocks = tuple(blocked_by.get(row[0], ()))
        flow.append(flow_issue(row, changes.get(row[0], []), projects, blocks,
                               versions_of(issue.get("fields") or {})))
    return Extract(rows, flow, links, len(items))
