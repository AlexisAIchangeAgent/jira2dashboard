"""The status rules that come from Jira (AL, 2026-11-18: the metrics use the _from_Jira sheets; the _from_user sheet Workflows only holds exceptions).

For every status (by name: a Jira status has one category in the instance) and for each workflow kind (Team, Feature) the rules the flow reads are:
Category (Backlog, Active, Wait, Done, Removed, Unclassified), IsStart, IsEnd, IsFree. What Jira says, by default:
  Category  Jira's status category: new -> Backlog, indeterminate (in progress) -> Active, done -> Done;
  IsEnd     1 for every Done status (so the person lists as exceptions the done statuses that are no delivery: Removed, Wait...);
  IsStart   the first in-progress status reached from the status of creation in the transitions of the workflow (first_in_progress, computed
            at the Jira read: it fills the column IsStart of the sheet JiraWorkflows);
  IsFree    1 for a status that a GLOBAL transition reaches (JiraWorkflows.IsFree).
A cell the person filled in the sheet Workflows wins, field by field (an empty cell takes Jira's value), a row that names a workflow only touches that
workflow. Jira cannot say Active or Wait, nor Removed: those stay exceptions, and Flow Efficiency needs at least one Wait status (has_wait).
The category Jira gives the CURRENT status of a ticket (Issues.StatusCategory) is the second source, for a status that the catalog does not list
(a status of a workflow that was retired); a status that no source knows stays Unclassified, named by the sheet Reconciliation.
VBA: modStatusRules, the same rules (the cases of tests/unit/test_status_rules.py are in modTests).
"""
from collections import deque
from dataclasses import dataclass
from typing import Iterable

from jira2dashboard.flow.model import ACTIVE, BACKLOG, DONE, UNCLASSIFIED, WAIT

_CATEGORY = {"new": BACKLOG, "todo": BACKLOG, "to do": BACKLOG, "indeterminate": ACTIVE, "in_progress": ACTIVE, "in progress": ACTIVE, "done": DONE}


def jira_category(key: str | None) -> str | None:
    """The category of the workbook for a status category of Jira (the key of the statuses call or the enum of the preview), None when unknown."""
    return _CATEGORY.get((key or "").strip().lower())


@dataclass(frozen=True)
class Exception_:
    """One row of the sheet Workflows: a cell left empty is None (take Jira's value); `workflow` None = both workflows."""
    workflow: str | None
    status: str
    category: str | None
    is_start: int | None
    is_end: int | None
    is_free: int | None


@dataclass(frozen=True)
class Rule:
    workflow: str
    status: str
    category: str
    is_start: int
    is_end: int
    is_free: int
    source: str                    # "user" when an exception of the sheet Workflows touches it, else "Jira"


def first_in_progress(initial: str | None, moves: Iterable[tuple[str, str]], categories: dict[str, str]) -> str | None:
    """The start of a workflow: the first status in progress (Jira category indeterminate) that a breadth-first search reaches from the status of
    creation along `moves` (from, to). The nearest wins, the first listed breaks a tie; None when the creation status is unknown or nothing is reached."""
    if not initial:
        return None
    graph: dict[str, list[str]] = {}
    for source, target in moves:
        graph.setdefault(source, []).append(target)
    seen, queue = {initial}, deque([initial])
    while queue:
        status = queue.popleft()
        if jira_category(categories.get(status)) == ACTIVE:
            return status
        for target in graph.get(status, ()):
            if target not in seen:
                seen.add(target)
                queue.append(target)
    return None


def flow_depths(initial: str | None, moves: Iterable[tuple[str, str]]) -> dict[str, int]:
    """The number of moves from the status of creation to each status that the transitions reach (breadth-first search): 0 for the creation status, 1
    for the next ones... It orders the statuses of a workflow as the normal flow goes (the stack of the chart Cycle time by status). A status that no
    transition reaches (a free one, reached from anywhere) has no depth. VBA twin: modParams.FlowDepths."""
    if not initial:
        return {}
    graph: dict[str, list[str]] = {}
    for source, target in moves:
        graph.setdefault(source, []).append(target)
    depth, queue = {initial: 0}, deque([initial])
    while queue:
        status = queue.popleft()
        for target in graph.get(status, ()):
            if target not in depth:
                depth[target] = depth[status] + 1
                queue.append(target)
    return depth


def _applying(exceptions: list[Exception_], kind: str, status: str) -> list[Exception_]:
    """The exceptions of a status for a workflow, the most precise last (a row of the workflow beats a row of both)."""
    mine = [e for e in exceptions if e.status == status and e.workflow in (None, kind)]
    return sorted(mine, key=lambda e: e.workflow is not None)


def effective_rules(kinds: Iterable[str], catalog: dict[str, str], current: dict[str, str], starts: set[str], free: set[str],
                    exceptions: list[Exception_]) -> list[Rule]:
    """One rule per workflow kind and status, the kinds in the given order then the names in alphabetical order. `catalog`: status -> Jira category
    (the statuses call), `current`: status -> Jira category of the tickets in it, `starts` / `free`: the statuses that Jira's workflows start with / reach
    freely."""
    statuses = sorted(set(catalog) | set(current) | {e.status for e in exceptions})
    out = []
    for kind in kinds:
        for status in statuses:
            category = jira_category(catalog.get(status)) or jira_category(current.get(status)) or UNCLASSIFIED
            is_start, is_free, is_end = int(status in starts), int(status in free), None
            applied = _applying(exceptions, kind, status)
            for e in applied:
                category = e.category or category
                is_start = is_start if e.is_start is None else e.is_start
                is_free = is_free if e.is_free is None else e.is_free
                is_end = e.is_end if e.is_end is not None else is_end
            if is_end is None:
                is_end = int(category == DONE)
            out.append(Rule(kind, status, category, is_start, is_end, is_free, "user" if applied else "Jira"))
    return out


def has_wait(rules: Iterable[Rule]) -> bool:
    """True once a status is declared Wait: without one Flow Efficiency would be 100 % by construction, so it is not computable."""
    return any(r.category == WAIT for r in rules)


# --- what the rules hold, and the warnings when a metric cannot be computed (AL, 2026-11-22: "flow metrics are still not computable") ----------
# The texts are the same in the VBA (modParams.RuleCountsLine, RuleWarnings): tests/unit/test_status_rules.py reads them in modParams.bas.

NO_START_CODE, NO_WAIT_CODE = "J2D-FLOW-010", "J2D-FLOW-011"
NO_START_TEXT = (" workflow: Flow Time, Flow Efficiency and Flow Load cannot be computed | declare IsStart = 1 for its first in-progress status in the "
                 "sheet Workflows")
NO_WAIT_TEXT = ("no status is declared Wait: Flow Efficiency cannot be computed | declare Category = Wait in the sheet Workflows for the statuses where "
                "the work waits")


@dataclass(frozen=True)
class KindCounts:
    """How many statuses of one workflow kind start the flow, are Active, Wait, and end it."""
    kind: str
    start: int
    active: int
    wait: int
    end: int


def rule_counts(rules: Iterable[Rule], kinds: Iterable[str]) -> list[KindCounts]:
    """The counts of each kind, in the order of `kinds`."""
    rules = list(rules)
    return [KindCounts(kind, sum(r.is_start == 1 for r in rules if r.workflow == kind), sum(r.category == ACTIVE for r in rules if r.workflow == kind),
                       sum(r.category == WAIT for r in rules if r.workflow == kind), sum(r.is_end == 1 for r in rules if r.workflow == kind))
            for kind in kinds]


def counts_line(counts: Iterable[KindCounts]) -> str:
    """The INFO line of the log: `status rules | Feature: start 1, active 2, wait 0, end 2 | Team: ...`."""
    return "status rules | " + " | ".join(f"{c.kind}: start {c.start}, active {c.active}, wait {c.wait}, end {c.end}" for c in counts)


def rule_warnings(counts: Iterable[KindCounts], kinds_with_tickets: set[str]) -> list[str]:
    """The WARN lines: a kind that has tickets and no start status (no StartUtc, so no Flow Time, Flow Efficiency or Flow Load), then no Wait status in
    any kind (Flow Efficiency reads the whole sheet StatusRules). Nothing when no kind has a ticket."""
    counts = list(counts)
    out = [f"{NO_START_CODE} | no start status in the {c.kind}{NO_START_TEXT}" for c in counts if c.kind in kinds_with_tickets and c.start == 0]
    if kinds_with_tickets and not any(c.wait for c in counts):
        out.append(f"{NO_WAIT_CODE} | {NO_WAIT_TEXT}")
    return out
