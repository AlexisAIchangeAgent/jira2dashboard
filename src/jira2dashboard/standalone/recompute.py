"""The raw sheets of a filled workbook read back into what a run read from Jira (AL, 2026-11-22, step F: `scripts\\run.ps1 compute`).

A workbook filled by the macros (button 1. Extract from Jira, then saved) or the .xlsx of the Python path holds what the extraction read: the
sheets Issues, Changelog and IssueLinks, and the Jira sheets of the configuration (JiraStatusCatalog, JiraIssueTypes, JiraWorkflows,
JiraTransitions). From them this module gives back the rows, the flow issues, the status catalog and the workflows, so that the metrics are
computed again without calling Jira: a corrected sheet Workflows, Projects or Config takes effect, as button 2. Compute the metrics does.
Never a derived sheet: the status intervals, their categories and the rules (StatusRules) are made again from the raw data and the parameters,
or a corrected exception would change nothing. Pure but for `read_workbook` (a zip read, no Excel).
"""
import datetime as dt
from dataclasses import dataclass
from pathlib import Path

from jira2dashboard import sheet_names as S
from jira2dashboard.flow import jira_workflows as JW
from jira2dashboard.flow.formulas import CHANGELOG_RAW, ISSUES
from jira2dashboard.flow.jira_workflows import JiraWorkflows
from jira2dashboard.flow.model import Change, FlowIssue, PIChange
from jira2dashboard.params.model import ParamsError
from jira2dashboard.standalone.mapping import date_shape, flow_issue, status_name
from jira2dashboard.standalone.xlsx_read import read_sheets

EPOCH = dt.datetime(1899, 12, 30)                    # Excel's day 0
ISSUE_DATES = ("CreatedUtc", "ResolvedUtc", "SprintStart", "SprintEnd", "TargetStart", "TargetEnd")
ISSUE_TEXTS = ("Key", "Id", "Project", "Type", "Status", "Parent", "StatusCategory", "FixVersions", "PI", "TeamValue", "SprintState")
ISSUE_WHOLE = ("SprintCount", "KeyChanges")
CHANGELOG_IDS = ("IssueId", "ChangeId")
RAW = ("Issues", "Changelog", "IssueLinks")
JIRA = ("JiraStatusCatalog", "JiraIssueTypes", "JiraWorkflows", "JiraTransitions")
NO_EXTRACTION = "J2D-FLOW-001"


def from_serial(value) -> dt.datetime | None:
    """An Excel date (a serial number) as the UTC moment the extraction wrote, to the second (both writers keep whole seconds)."""
    if value is None or value == "":
        return None
    if isinstance(value, dt.datetime):
        return value
    return EPOCH + dt.timedelta(seconds=round(float(value) * 86400))


def as_text(value) -> str | None:
    """A text cell; an id that a file stored as a number (10001.0) gives back its digits. Empty is None."""
    if value is None or value == "":
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _columns(grid: list[list], header: tuple[str, ...], sheet: str) -> list[dict]:
    """The rows of a sheet as dicts by the names of `header` (row 1 holds the names, in any order); the one blank row of an empty table is no row."""
    names = [None if c is None else str(c) for c in (grid[0] if grid else [])]
    missing = [h for h in header if h not in names]
    if missing:
        raise ParamsError(f"The sheet {S.of(sheet)} has no column {', '.join(missing)}: run 1. Extract from Jira again", "J2D-FLOW-003")
    at = {h: names.index(h) for h in header}
    rows = []
    for row in grid[1:]:
        cells = {h: row[at[h]] if at[h] < len(row) else None for h in header}
        if any(v not in (None, "") for v in cells.values()):
            rows.append(cells)
    return rows


def issue_rows(grid: list[list]) -> list[tuple]:
    """The sheet Issues as the rows `mapping.issue_row` gives (formulas.ISSUES): dates, texts and whole numbers restored."""
    out = []
    for cells in _columns(grid, ISSUES, "Issues"):
        for name in ISSUE_DATES:
            cells[name] = from_serial(cells[name])
        for name in ISSUE_TEXTS:
            cells[name] = as_text(cells[name])
        for name in ISSUE_WHOLE:
            cells[name] = int(cells[name] or 0)
        cells["Id"] = cells["Id"] or ""
        out.append(tuple(cells[name] for name in ISSUES))
    return out


def changelog_rows(grid: list[list]) -> list[tuple]:
    """The sheet Changelog (formulas.CHANGELOG_RAW) with its date and its ids restored, in the order of the sheet (that of Jira)."""
    out = []
    for cells in _columns(grid, CHANGELOG_RAW, "Changelog"):
        cells["CreatedUtc"] = from_serial(cells["CreatedUtc"])
        for name in CHANGELOG_IDS:
            cells[name] = as_text(cells[name])
        out.append(tuple(cells[name] for name in CHANGELOG_RAW))
    return out


def unreadable(changelog: list[tuple]) -> tuple[int, str]:
    """The status changes whose date could not be read, with their shapes (mapping.unreadable_status_dates, from the sheet: CreatedText)."""
    shapes: dict[str, int] = {}
    for row in changelog:
        if row[5] == "status" and row[3] is None:
            shape = date_shape(row[4])
            shapes[shape] = shapes.get(shape, 0) + 1
    return sum(shapes.values()), "; ".join(f"{shape} x {count}" for shape, count in shapes.items())


def flow_of(rows: list[tuple], changelog: list[tuple], links: list[tuple[str, str]], projects: list[tuple[str, str, str]], pi_field: str,
            feature_types: frozenset[str] | None, names: dict[str, str] | None = None) -> list[FlowIssue]:
    """The flow issues of the rows of Issues, as `mapping.extracted` makes them: the status changes and the changes of the PI field (Config.PIField)
    of the sheet Changelog (a change whose date could not be read is left out), the blocking links, the fix versions of the row. `names`: status id ->
    its name now (JiraStatusCatalog): the sheet Changelog keeps the text Jira sent, a renamed status takes its name of now (mapping.status_name)."""
    changes: dict[str, list[Change]] = {}
    pi_history: dict[str, list[PIChange]] = {}
    for key, _issue, change, at, _text, field_id, _field, _type, from_id, from_string, to_id, to_string in changelog:
        if at is None:
            continue
        if field_id == "status":
            changes.setdefault(key, []).append(Change(change or "", at, status_name(as_text(from_id), from_string, names),
                                                      status_name(as_text(to_id), to_string, names)))
        elif pi_field and field_id == pi_field:
            pi_history.setdefault(key, []).append(PIChange(change or "", at, from_string, to_string))
    blocked_by: dict[str, list[str]] = {}
    for blocker, blocked in links:
        blocked_by.setdefault(blocker, []).append(blocked)
    return [flow_issue(row, changes.get(row[0], []), projects, tuple(blocked_by.get(row[0], ())),
                       tuple(v.strip() for v in row[10].split(",")) if row[10] else (), pi_history.get(row[0], []), feature_types)
            for row in rows]


@dataclass(frozen=True)
class Filled:
    """What the extraction left in the workbook."""
    rows: list[tuple]                   # Issues
    changelog: list[tuple]              # Changelog
    links: list[tuple[str, str]]        # IssueLinks
    catalog: dict[str, str]             # JiraStatusCatalog: status -> Jira category
    names: dict[str, str]               # JiraStatusCatalog: status id -> name (a renamed status of the changelog takes its name of now)
    workflows: JiraWorkflows            # JiraIssueTypes, JiraWorkflows, JiraTransitions (empty when the workflows were not read)


def filled(grids: dict[str, list[list]]) -> Filled:
    """`grids` by LOGICAL sheet name. The raw sheets must be there and Issues must hold a ticket; a Jira sheet missing or empty gives nothing."""
    missing = [name for name in RAW if name not in grids]
    if missing:
        raise ParamsError(f"The workbook has no sheet {', '.join(S.of(n) for n in missing)}: run 1. Extract from Jira first", NO_EXTRACTION)
    rows = issue_rows(grids["Issues"])
    if not rows:
        raise ParamsError(f"The sheet {S.of('Issues')} holds no ticket: run 1. Extract from Jira first", NO_EXTRACTION)
    links = [(as_text(c["Blocker"]), as_text(c["Blocked"])) for c in _columns(grids["IssueLinks"], ("Blocker", "Blocked"), "IssueLinks")]
    body = {name: [tuple(r) for r in grids.get(name, [])[1:] if any(v not in (None, "") for v in r)] for name in JIRA}
    return Filled(rows, changelog_rows(grids["Changelog"]), links, JW.catalog_of(body["JiraStatusCatalog"]), JW.names_of(body["JiraStatusCatalog"]),
                  JW.from_sheets(body["JiraIssueTypes"], body["JiraWorkflows"], body["JiraTransitions"]))


def read_workbook(path: Path) -> Filled:
    """The raw and the Jira sheets of the workbook at `path` (an .xlsm saved by Excel, or the .xlsx of the Python path), read as a zip."""
    names = {S.of(name): name for name in RAW + JIRA}
    return filled({names[real]: grid for real, grid in read_sheets(path, set(names)).items()})
