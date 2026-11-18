"""Domain vocabulary of the flow metrics (names as in the documentation). All datetimes are naive UTC."""
import datetime as dt
from dataclasses import dataclass

from jira2dashboard.params.model import ParamSheet, load_all

NOT_COMPUTABLE = "not computable"        # shown instead of a figure that cannot be computed: never 0, never a guess
BACKLOG, ACTIVE, WAIT, DONE, REMOVED, UNCLASSIFIED = "Backlog", "Active", "Wait", "Done", "Removed", "Unclassified"
WORK_TYPES = ("Feature", "Enabler", "Defect", "Autre")        # TypeCategorie; anything else counts as Unmapped
UNMAPPED = "Unmapped"
ART = "ART"                                                    # team name under which Program features are reported
OTHER_TYPES = "ART (other types)"                              # the team of the Program tickets that are not Features (sub-tasks...): no ART metric counts them
MODES = ("A", "B")                                             # FlowLoadMode


@dataclass(frozen=True)
class StatusRule:
    """One row of the Workflows sheet."""
    category: str
    is_start: bool
    is_end: bool
    is_free: bool


class Rules:
    """The status rules: what Jira says (flow/status_rules.py) and the exceptions of the sheet Workflows. A status that no source knows is
    Unclassified. `waits` is True once a status is declared Wait (Flow Efficiency needs one)."""

    def __init__(self, rows: dict[tuple[str, str], StatusRule], waits: bool = True):
        self._rows = rows
        self.waits = waits

    @classmethod
    def from_params(cls, sheets: list[ParamSheet] | None = None, catalog: dict[str, str] | None = None, current: dict[str, str] | None = None,
                    starts: set[str] | None = None, free: set[str] | None = None) -> "Rules":
        """The rules of the sheet Workflows (its rows are exceptions) over what Jira gives: `catalog` status -> Jira category, `current` the
        category Jira gives the tickets' current statuses, `starts` / `free` the statuses that the Jira workflows start with / reach freely.
        Without any Jira argument the rows alone rule (a complete sheet gives the rules of before)."""
        from jira2dashboard.flow.status_rules import Exception_, effective_rules, has_wait
        workflows = next(s for s in (sheets or load_all()) if s.name == "Workflows")
        exceptions = [Exception_(wf, status, category, start, end, free_) for wf, status, category, start, end, free_ in workflows.rows]
        rules = effective_rules(("Team", "Feature"), catalog or {}, current or {}, set(starts or ()), set(free or ()), exceptions)
        return cls({(r.workflow, r.status): StatusRule(r.category, bool(r.is_start), bool(r.is_end), bool(r.is_free)) for r in rules},
                   waits=has_wait(rules))

    def ended_statuses(self) -> set[str]:
        """The statuses that end the flow (IsEnd) in at least one workflow: what the sheet Reconciliation lists with the tickets that sit in them."""
        return {status for (_wf, status), rule in self._rows.items() if rule.is_end}

    def rule(self, workflow: str, status: str) -> StatusRule:
        return self._rows.get((workflow, status), StatusRule(UNCLASSIFIED, False, False, False))


@dataclass(frozen=True)
class Change:
    """A status change of an issue."""
    id: str
    at: dt.datetime
    from_status: str
    to_status: str


@dataclass(frozen=True)
class PIChange:
    """A change of the PI field of an issue: the changelog item whose fieldId is Config.PIField, texts as Jira sent them
    (None when the field was empty before or was emptied)."""
    id: str
    at: dt.datetime
    from_value: str | None
    to_value: str | None


def change_order(change: Change) -> tuple:
    """The order of the changes of an issue: by date, and for two changes in the same second (the milliseconds are dropped) by
    their id as a number (9 before 10: a text sort would put 10 first), which is the order Jira gave them in (the VBA keeps
    it: modFlow.SortByDate is stable)."""
    return (change.at, len(change.id), change.id)


def program_team(type_: str, feature_types: frozenset[str] | None) -> str:
    """The team of a ticket of the Program project: ART for a Feature (its type is in Config.FeatureTypes), OTHER_TYPES for the rest (AL, 2026-11-18:
    for a Program the flow metrics look at the Features). `feature_types` None = no filter, every type is a Feature (a caller that does not know the
    parameter); an empty set = no Feature at all, as the workbook's `IsFeature` column does."""
    return ART if feature_types is None or type_ in feature_types else OTHER_TYPES


@dataclass(frozen=True)
class FlowIssue:
    """What the metrics need to know about an issue (the extract plus its status changes)."""
    key: str
    team: str                      # team name, or ART for Program features
    type: str                      # issue type: Feature, Story, Task, Bug, Sub-task
    work_type: str | None          # TypeCategorie
    workflow: str                  # Team or Feature
    created: dt.datetime
    status: str                    # current status
    changes: tuple[Change, ...] = ()
    parent: str | None = None      # key of the parent issue (data check DC01)
    resolved: dt.datetime | None = None            # resolution date (DC04, DC05)
    status_category: str = ""      # Jira status category key: new, indeterminate or done (DC04, DC05)
    fix_versions: tuple[str, ...] = ()
    story_points: float | None = None
    pi: str | None = None          # value of the PI field, as typed
    business_value: float | None = None
    team_value: str | None = None  # value of the Team field
    sprint_state: str = ""         # state of the current sprint: active, else future, else closed; "" = no sprint
    sprint_start: dt.datetime | None = None        # dates of that sprint
    sprint_end: dt.datetime | None = None
    sprint_count: int = 0          # number of sprints in the Sprint field
    target_start: dt.datetime | None = None
    target_end: dt.datetime | None = None
    blocks: tuple[str, ...] = ()   # keys of the issues that this issue blocks
    key_changes: int = 0           # moves between projects
    pi_changes: tuple[PIChange, ...] = ()          # history of the PI field (what the PI was before each change)
    jira_workflow: str = ""         # the Jira workflow of its project and issue type (JiraIssueTypes); "" = not read (DC28)


@dataclass(frozen=True)
class Interval:
    """One visit of a status (sheet StatusIntervals)."""
    key: str
    seq: int
    status: str
    category: str
    entered: dt.datetime
    left: dt.datetime | None       # None while the issue is still there
    duration_days: float
    is_open: bool
    change_id: str | None          # the change that ended it


@dataclass(frozen=True)
class Iteration:
    """An iteration or a PI. Half-open: start <= date < end."""
    name: str
    start: dt.datetime
    end: dt.datetime

    def contains(self, moment: dt.datetime | None) -> bool:
        return moment is not None and self.start <= moment < self.end


def days_between(start: dt.datetime, end: dt.datetime) -> float:
    return (end - start).total_seconds() / 86400
