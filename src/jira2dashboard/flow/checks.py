"""Data checks that the extraction can make: DC01, DC04 to DC10 and DC28 (docs/safe-flow-metrics-and-data-checks.md, section 3).

Pure. The Excel formulas (formulas.py, columns DC.. of ItemFlow and sheet DataChecks) are tested against `issue_checks`,
the reference written here from the issue fields and the status history. The other checks (DC02, DC03, DC11 to DC27, DC29)
need Jira fields that the extraction does not read yet, or a list of feature types.

Choices (proposals, the document says only "StatusIntervals[Check] contains it"):
* DC01 applies to the issues of the Team workflow whose type is in `Config.ParentRequiredTypes` (default Story, Task).
* DC04 and DC05 use the Jira status category (key `done`) and the resolution date.
* DC06 covers FROM_MISMATCH (a change starts from a status that is not the one the issue was in) and CURRENT_MISMATCH (the
  history does not end in the status the issue has now): both say "the status history is inconsistent".
* DC28 compares the status of two consecutive intervals, so it follows what the intervals show, with Transitions and IsFree.
"""
import datetime as dt
from dataclasses import dataclass

from jira2dashboard.flow.itemflow import item_flow
from jira2dashboard.flow.model import UNCLASSIFIED, FlowIssue, Rules
from jira2dashboard.params.model import ParamSheet, load_all


@dataclass(frozen=True)
class Check:
    id: str
    title: str
    severity: str


CHECKS = (
    Check("DC01", "Story or task without parent", "Blocker"),
    Check("DC04", "Done without resolution date", "Major"),
    Check("DC05", "Resolution date but not done", "Major"),
    Check("DC06", "Status history inconsistent (FROM_MISMATCH or CURRENT_MISMATCH)", "Major"),
    Check("DC07", "Status change dated before the creation of the issue (BEFORE_CREATED)", "Major"),
    Check("DC08", "Status not in Workflows (category Unclassified)", "Blocker for efficiency"),
    Check("DC09", "Done without ever being in a start status", "Major"),
    Check("DC10", "Reopened after Done", "Major"),
    Check("DC28", "Status transition not in Transitions", "Major"),
)
IDS = tuple(c.id for c in CHECKS)


class Moves:
    """The allowed moves (sheet Transitions): (workflow, from status, to status)."""

    def __init__(self, rows: set[tuple[str, str, str]]):
        self._rows = rows

    @classmethod
    def from_params(cls, sheets: list[ParamSheet] | None = None) -> "Moves":
        transitions = next(s for s in (sheets or load_all()) if s.name == "Transitions")
        return cls({(wf, source, target) for wf, source, target, _kind in transitions.rows})

    def allowed(self, workflow: str, source: str, target: str) -> bool:
        return (workflow, source, target) in self._rows


def parent_required_types(sheets: list[ParamSheet] | None = None) -> frozenset[str]:
    """Config.ParentRequiredTypes: the issue types that must have a parent (comma separated)."""
    config = next(s for s in (sheets or load_all()) if s.name == "Config")
    value = next((row[1] for row in config.rows if row[0] == "ParentRequiredTypes"), None) or ""
    return frozenset(t.strip() for t in str(value).split(",") if t.strip())


def issue_checks(issue: FlowIssue, as_of: dt.datetime, rules: Rules, moves: Moves,
                 parent_types: frozenset[str]) -> frozenset[str]:
    """The ids of the checks that this issue fails."""
    found: set[str] = set()
    done = issue.status_category == "done"
    if issue.workflow == "Team" and issue.type in parent_types and not issue.parent:
        found.add("DC01")
    if done and issue.resolved is None:
        found.add("DC04")
    if not done and issue.resolved is not None:
        found.add("DC05")
    changes = sorted(issue.changes, key=lambda c: (c.at, c.id))
    status = changes[0].from_status if changes else issue.status
    entered = issue.created
    for change in changes:
        if change.from_status != status:
            found.add("DC06")
        if change.at < entered:
            found.add("DC07")
        status, entered = change.to_status, change.at
    if status != issue.status:
        found.add("DC06")
    item = item_flow(issue, as_of, rules)
    statuses = [i.status for i in item.intervals]
    if any(rules.rule(issue.workflow, s).category == UNCLASSIFIED for s in statuses):
        found.add("DC08")
    if item.skipped_start:
        found.add("DC09")
    if item.reopened > 0:
        found.add("DC10")
    for before, after in zip(statuses, statuses[1:]):
        if not moves.allowed(issue.workflow, before, after) and not rules.rule(issue.workflow, after).is_free:
            found.add("DC28")
    return frozenset(found)
