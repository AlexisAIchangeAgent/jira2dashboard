"""Domain vocabulary of the flow metrics (names as in the documentation). All datetimes are naive UTC."""
import datetime as dt
from dataclasses import dataclass

from jira2dashboard.params.model import ParamSheet, load_all

NOT_COMPUTABLE = "not computable"        # shown instead of a figure that cannot be computed: never 0, never a guess
BACKLOG, ACTIVE, WAIT, DONE, REMOVED, UNCLASSIFIED = "Backlog", "Active", "Wait", "Done", "Removed", "Unclassified"
WORK_TYPES = ("Feature", "Enabler", "Defect", "Autre")        # TypeCategorie; anything else counts as Unmapped
UNMAPPED = "Unmapped"
ART = "ART"                                                    # team name under which Program features are reported
MODES = ("A", "B")                                             # FlowLoadMode


@dataclass(frozen=True)
class StatusRule:
    """One row of the Workflows sheet."""
    category: str
    is_start: bool
    is_end: bool
    is_free: bool


class Rules:
    """The status mapping (sheet Workflows). A status that is not listed is Unclassified."""

    def __init__(self, rows: dict[tuple[str, str], StatusRule]):
        self._rows = rows

    @classmethod
    def from_params(cls, sheets: list[ParamSheet] | None = None) -> "Rules":
        workflows = next(s for s in (sheets or load_all()) if s.name == "Workflows")
        return cls({(wf, status): StatusRule(category, bool(start), bool(end), bool(free))
                    for wf, _order, status, category, start, end, free in workflows.rows})

    def rule(self, workflow: str, status: str) -> StatusRule:
        return self._rows.get((workflow, status), StatusRule(UNCLASSIFIED, False, False, False))


@dataclass(frozen=True)
class Change:
    """A status change of an issue."""
    id: str
    at: dt.datetime
    from_status: str
    to_status: str


def change_order(change: Change) -> tuple:
    """The order of the changes of an issue: by date, and for two changes in the same second (the milliseconds are dropped) by
    their id as a number (9 before 10: a text sort would put 10 first), which is the order Jira gave them in (the VBA keeps
    it: modFlow.SortByDate is stable)."""
    return (change.at, len(change.id), change.id)


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
