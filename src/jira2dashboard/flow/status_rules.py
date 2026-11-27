"""The status rules that come from Jira (AL, 2026-11-18: the metrics use the _from_Jira sheets; the _from_user sheet Workflows only holds exceptions).

For every status and for each workflow kind (Team, Feature) the rules the flow reads are: Category (Backlog, Active, Wait, Done, Removed, Unclassified),
IsStart, IsEnd, IsFree. A status is known by its KEY (model.status_key: the name trimmed, case ignored), the way Excel compares texts: two spellings of
one name ("En Analyse", "En analyse": two projects of the site) are ONE rule, or the SUMIFS of the formulas would add their flags up (AL, 2026-11-27:
IsStart 2, so no Feature of the client ever entered its start status). What Jira says, by default:
  Category  Jira's status category: new -> Backlog, indeterminate (in progress) -> Active, done -> Done;
  IsEnd     1 for every Done status (so the person lists as exceptions the done statuses that are no delivery: Removed, Wait...);
  IsStart   1 for every status in progress (its category, after the exceptions, is Active or Wait): the flow starts at the first entry into a status
            in progress after the last entry into a Backlog one, whatever the workflow, its version or the names of its statuses (AL, 2026-11-27:
            "the calculations must still work when the workflow changes"). Config FeatureStartStatuses / TeamStartStatuses, when filled, replace
            this for their kind: only the statuses they name start the flow;
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

from jira2dashboard.flow.model import ACTIVE, BACKLOG, DONE, UNCLASSIFIED, WAIT, status_key

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


IN_PROGRESS = (ACTIVE, WAIT)          # the categories that start the flow, unless Config names the start statuses of the kind
START_PARAMETERS = {"Feature": "FeatureStartStatuses", "Team": "TeamStartStatuses"}      # Config: the start statuses of a kind (a list), empty = in progress


def _applying(exceptions: list[Exception_], kind: str, key: str) -> list[Exception_]:
    """The exceptions of a status (by its key) for a workflow, the most precise last (a row of the workflow beats a row of both)."""
    mine = [e for e in exceptions if status_key(e.status) == key and e.workflow in (None, kind)]
    return sorted(mine, key=lambda e: e.workflow is not None)


def _category(names: list[str], source: dict[str, str]) -> str | None:
    """The category of the first spelling (in the order of `names`) that `source` knows."""
    return next((c for c in (jira_category(source.get(n)) for n in names) if c), None)


def spellings(catalog: Iterable[str], current: Iterable[str], exceptions: Iterable[Exception_]) -> dict[str, list[str]]:
    """key -> the spellings of the status: those of the catalog (binary order), then of the current statuses (binary order), then of the exceptions
    (their order, trimmed). The first one is the name of the rule."""
    out: dict[str, list[str]] = {}
    for source in (sorted(catalog), sorted(current), [e.status.strip() for e in exceptions]):
        for name in source:
            key = status_key(name)
            if key and name not in out.setdefault(key, []):
                out[key].append(name)
    return out


def effective_rules(kinds: Iterable[str], catalog: dict[str, str], current: dict[str, str], free: set[str], exceptions: list[Exception_],
                    start_statuses: dict[str, Iterable[str]] | None = None) -> list[Rule]:
    """One rule per workflow kind and status key, the kinds in the given order then the names in binary order. `catalog`: status -> Jira category (the
    statuses call), `current`: status -> Jira category of the tickets in it, `free`: the statuses that Jira's workflows reach freely,
    `start_statuses`: kind -> the names of Config FeatureStartStatuses / TeamStartStatuses (none: the statuses in progress start the flow).
    The Jira category of a key: the first spelling the catalog knows, else the first the current statuses know (`spellings`)."""
    names_of = spellings(catalog, current, exceptions)
    free_keys = {status_key(n) for n in free}
    named = {kind: {status_key(n) for n in names} - {""} for kind, names in (start_statuses or {}).items()}
    out = []
    for kind in kinds:
        for key, names in sorted(names_of.items(), key=lambda kv: kv[1][0]):
            category = _category(names, catalog) or _category(names, current) or UNCLASSIFIED
            is_free, is_end = int(key in free_keys), None
            applied = _applying(exceptions, kind, key)
            for e in applied:
                category = e.category or category
                is_free = is_free if e.is_free is None else e.is_free
                is_end = e.is_end if e.is_end is not None else is_end
            if is_end is None:
                is_end = int(category == DONE)
            is_start = int(key in named[kind]) if named.get(kind) else int(category in IN_PROGRESS)
            out.append(Rule(kind, names[0], category, is_start, is_end, is_free, "user" if applied else "Jira"))
    return out


def unknown_statuses(exceptions: Iterable[Exception_], start_statuses: dict[str, Iterable[str]], known: Iterable[str]) -> list[tuple[str, str]]:
    """The statuses that a setting names and no source knows (`known`: the catalog, the current statuses and the histories), as (name, where) in the
    order of the sheet Workflows then of the kinds, each key once: an exception typed for a status that Jira spells otherwise ("À déployer" for
    "A déployer") rules nothing, and a status renamed in Jira leaves its old name here. `where` = "Workflows" or the Config parameter."""
    keys = {status_key(n) for n in known}
    named = [(e.status, "Workflows") for e in exceptions]
    named += [(n, START_PARAMETERS[kind]) for kind in START_PARAMETERS for n in (start_statuses or {}).get(kind, ())]
    out, seen = [], set()
    for name, where in named:
        key = status_key(name)
        if key and key not in keys and key not in seen:
            seen.add(key)
            out.append((str(name).strip(), where))
    return out


def has_wait(rules: Iterable[Rule]) -> bool:
    """True once a status is declared Wait: without one Flow Efficiency would be 100 % by construction, so it is not computable."""
    return any(r.category == WAIT for r in rules)


# --- what the rules hold, and the warnings when a metric cannot be computed (AL, 2026-11-22: "flow metrics are still not computable") ----------
# The texts are the same in the VBA (modParams.RuleCountsLine, RuleWarnings): tests/unit/test_status_rules.py reads them in modParams.bas.

NO_START_CODE, NO_WAIT_CODE = "J2D-FLOW-010", "J2D-FLOW-011"
NO_START_TEXT = (" workflow: Flow Time, Flow Efficiency and Flow Load cannot be computed | the statuses in progress start the flow: check the "
                 "Category of its statuses in the sheet Workflows_from_user and its start statuses in the sheet Config_from_user")
NO_WAIT_TEXT = ("no status is declared Wait: Flow Efficiency cannot be computed | declare Category = Wait in the sheet Workflows_from_user for the statuses where "
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


STATUS_RULES_HEADER = ("Workflow", "Status", "Category", "IsStart", "IsEnd", "IsFree", "Source", "JiraWorkflows")    # modExtract (both writers)


def status_rule_rows(rules: Iterable[Rule], names: dict[tuple[str, str], str] | None = None) -> list[tuple]:
    """The sheet StatusRules: the Feature rules then the Team ones, by status (the order of modParams.EffectiveRules(Array("Feature", "Team"))); the
    last column, the Jira workflows that hold the status (jira_workflows.jira_workflow_names)."""
    names = names or {}
    return [(r.workflow, r.status, r.category, r.is_start, r.is_end, r.is_free, r.source, names.get((r.workflow, r.status), ""))
            for r in sorted(rules, key=lambda r: (r.workflow != "Feature", r.status))]
