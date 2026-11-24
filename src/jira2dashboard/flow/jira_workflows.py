"""What Jira says about workflows (AL, 2026-11-18): the workflow of each issue type of each project, its creation status, its moves, the statuses that start
it and those reached from anywhere. The sheets JiraIssueTypes, JiraWorkflows and JiraTransitions hold it in the workbook; the Python path reads it from the
same calls (standalone/workflows.py) and the fake Jira's definition gives it to the tests (mockjira/flow_adapter.jira_workflows). Pure.

It feeds: the start and the free statuses of the status rules (status_rules), the order of the stack of the chart Cycle time by status (feature_depths),
and the check DC28 (a move that the workflow of the ticket does not allow).
"""
from dataclasses import dataclass, field
from typing import Iterable

from jira2dashboard.flow.status_rules import flow_depths


@dataclass(frozen=True)
class JiraWorkflows:
    initial: dict[str, str] = field(default_factory=dict)                       # workflow -> its creation status
    moves: frozenset[tuple[str, str, str]] = frozenset()                       # (workflow, from status, to status)
    starts: frozenset[str] = frozenset()                                       # the statuses that start a workflow (first in progress from creation)
    free: frozenset[str] = frozenset()                                         # the statuses a GLOBAL transition reaches
    workflow_of: dict[tuple[str, str], str] = field(default_factory=dict)      # (project, issue type) -> workflow name

    def knows(self, workflow: str) -> bool:
        """True when the moves of this workflow were read: only then can a move be judged."""
        return any(wf == workflow for wf, _source, _target in self.moves)

    def allowed(self, workflow: str, source: str, target: str) -> bool:
        return (workflow, source, target) in self.moves

    def feature_depths(self, programs: Iterable[str], feature_types: Iterable[str]) -> dict[str, int]:
        """status -> its smallest depth (moves from the creation status) in the workflows the Feature types of the Program projects follow."""
        names = {self.workflow_of[(project, kind)] for project in programs for kind in feature_types if (project, kind) in self.workflow_of}
        depths: dict[str, int] = {}
        for name in sorted(names):
            moves = [(source, target) for wf, source, target in self.moves if wf == name]
            for status, depth in flow_depths(self.initial.get(name), moves).items():
                depths[status] = min(depth, depths.get(status, depth))
        return depths


# --- the sheets of the workbook (AL, 2026-11-24, step F): the layout of the VBA (modJiraConfig.WriteWorkflowSheets, WriteStatusCatalog; a test reads
# the .bas), so that a sheet written by either side is read by both. The readers take only the columns the VBA reads back: JiraIssueTypes A..C and
# JiraTransitions A..C (the formulas), JiraWorkflows Status | IsStart | IsFree | Initial (FlaggedStatuses, StatusNames), JiraStatusCatalog Status |
# JiraCategory (RefreshRules). What the Python reading does not know stays empty (the Id of a status, the Category, IsEnd and Project of a row).
CATALOG_HEADER = ("Id", "Status", "JiraCategory")
ISSUE_TYPES_HEADER = ("Project", "IssueType", "Workflow", "Subtask", "HierarchyLevel", "ProjectStyle")
WORKFLOWS_HEADER = ("Workflow", "Status", "Category", "IsStart", "IsEnd", "IsFree", "JiraCategory", "Initial", "Project")
TRANSITIONS_HEADER = ("Workflow", "From", "To", "JiraTransition", "TransitionId", "Scope", "Project")


def catalog_rows(catalog: dict[str, str]) -> list[tuple]:
    """The sheet JiraStatusCatalog: one row per status of the instance, by name (the Python reading keeps no id)."""
    return [("", status, catalog[status]) for status in sorted(catalog)]


def catalog_of(rows: list[tuple]) -> dict[str, str]:
    """status -> Jira category, from the rows of JiraStatusCatalog (Id, Status, JiraCategory)."""
    return {str(r[1]): "" if r[2] is None else str(r[2]) for r in rows if r[1] not in (None, "")}


def sheet_rows(info: JiraWorkflows) -> dict[str, list[tuple]]:
    """The rows of JiraIssueTypes, JiraWorkflows and JiraTransitions. A workflow lists its creation status, then the statuses of its moves by name; a
    start or free status that no workflow lists gets a row of its own with no workflow (FlaggedStatuses reads every row)."""
    types = [(project, kind, info.workflow_of[(project, kind)], "", "", "") for project, kind in sorted(info.workflow_of)]
    flows, listed = [], set()
    for name in sorted(set(info.initial) | {wf for wf, _s, _t in info.moves}):
        first = info.initial.get(name)
        others = sorted({s for wf, a, b in info.moves if wf == name for s in (a, b)} - {first})
        for status in ([first] if first else []) + others:
            flows.append((name, status, "", 1 if status in info.starts else "", "", 1 if status in info.free else 0, "", 1 if status == first else "", ""))
            listed.add(status)
    flows += [("", status, "", 1 if status in info.starts else "", "", 1 if status in info.free else 0, "", "", "")
              for status in sorted((info.starts | info.free) - listed)]
    moves = [(wf, source, target, "", "", "DIRECTED", "") for wf, source, target in sorted(info.moves)]
    return {"JiraIssueTypes": types, "JiraWorkflows": flows, "JiraTransitions": moves}


def _flag(value) -> bool:
    try:
        return float(value) == 1
    except (TypeError, ValueError):
        return False


def from_sheets(issue_types: list[tuple], workflows: list[tuple], transitions: list[tuple]) -> JiraWorkflows:
    """What the run read of Jira's workflows, from the three sheets (written by the VBA or by sheet_rows). An empty Workflow is no workflow."""
    text = lambda v: "" if v is None else str(v)                                   # noqa: E731
    workflow_of = {(text(r[0]), text(r[1])): text(r[2]) for r in issue_types if text(r[2])}
    initial = {text(r[0]): text(r[1]) for r in workflows if text(r[0]) and _flag(r[7] if len(r) > 7 else None)}
    starts = frozenset(text(r[1]) for r in workflows if _flag(r[3]))
    free = frozenset(text(r[1]) for r in workflows if _flag(r[5]))
    moves = frozenset((text(r[0]), text(r[1]), text(r[2])) for r in transitions if text(r[0]))
    return JiraWorkflows(initial, moves, starts, free, workflow_of)
