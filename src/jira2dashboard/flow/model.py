"""Domain vocabulary of the flow metrics (names as in the documentation). All datetimes are naive UTC."""
import datetime as dt
import re
from dataclasses import dataclass

from jira2dashboard.params.model import ParamSheet, load_all

NOT_COMPUTABLE = "not computable"        # shown instead of a figure that cannot be computed: never 0, never a guess
BACKLOG, ACTIVE, WAIT, DONE, REMOVED, UNCLASSIFIED = "Backlog", "Active", "Wait", "Done", "Removed", "Unclassified"
WORK_TYPES = ("Feature", "Enabler", "Defect", "Autre")        # TypeCategorie; anything else counts as Unmapped
UNMAPPED = "Unmapped"
ART = "ART"                                                    # team name under which Program features are reported
OTHER_TYPES = "ART (other types)"                              # the team of the Program tickets that are not Features (sub-tasks...): no ART metric counts them
MODES = ("A", "B")                                             # FlowLoadMode


def work_type_bucket(work_type) -> str:
    """The bucket of Flow Distribution: a work type of WORK_TYPES whatever its case ("feature" is Feature, as the = of Excel reads it), its
    canonical name; anything else Unmapped (the parity check of 2026-11-23: Excel counted "feature" as Feature, Python as Unmapped)."""
    # a number or a boolean (WorkTypeField on a field of another type) is Unmapped, as Excel's 5="Feature" is FALSE
    return {w.lower(): w for w in WORK_TYPES}.get("" if work_type is None else str(work_type).lower(), UNMAPPED)


def status_key(name) -> str:
    """The key of a status name, the way Excel compares two texts (SUMIFS, COUNTIFS, MATCH): trimmed, case ignored. Two spellings of one name are
    one status for the formulas, so they must be one for the rules too (status_rules; VBA modParams.StatusKey: LCase$(TrimWhite(...)))."""
    return str(name or "").strip().lower()


def list_value(value) -> tuple[str, ...]:
    """The names of a list parameter of Config (FeatureTypes, FeatureStartStatuses...): separated by commas, trimmed, empty ones left out."""
    return tuple(t.strip() for t in str(value or "").split(",") if t.strip())


@dataclass(frozen=True)
class StatusRule:
    """One row of the Workflows sheet."""
    category: str
    is_start: bool
    is_end: bool
    is_free: bool


class Rules:
    """The status rules: what Jira says (flow/status_rules.py) and the exceptions of the sheet Workflows, keyed by (workflow, status_key). A status
    that no source knows is Unclassified. `waits` is True once a status is declared Wait (Flow Efficiency needs one)."""

    def __init__(self, rows: dict[tuple[str, str], StatusRule], waits: bool = True, entries: tuple = (), exceptions: tuple = (),
                 start_statuses: dict | None = None, known: frozenset = frozenset()):
        self._rows = {(wf, status_key(status)): rule for (wf, status), rule in rows.items()}
        self._names = {(wf, status_key(status)): status for (wf, status) in rows}
        self.waits = waits
        self.entries = entries          # the status_rules.Rule of each row with its source (from_params): the sheet StatusRules
        self.exceptions = exceptions    # the rows of the sheet Workflows (status_rules.Exception_)
        self.start_statuses = start_statuses or {}      # kind -> the names of Config FeatureStartStatuses / TeamStartStatuses
        self.known = known              # the statuses Jira knows (catalog and current statuses): what the health field "Statuses unknown" compares with

    @classmethod
    def from_params(cls, sheets: list[ParamSheet] | None = None, catalog: dict[str, str] | None = None, current: dict[str, str] | None = None,
                    free: set[str] | None = None) -> "Rules":
        """The rules of the sheet Workflows (its rows are exceptions) over what Jira gives: `catalog` status -> Jira category, `current` the
        category Jira gives the tickets' current statuses, `free` the statuses that the Jira workflows reach freely; the start statuses of each kind
        from Config (FeatureStartStatuses, TeamStartStatuses) when the sheets hold it. Without any Jira argument the rows alone rule."""
        from jira2dashboard.flow.status_rules import START_PARAMETERS, Exception_, effective_rules, has_wait
        sheets = sheets or load_all()
        workflows = next(s for s in sheets if s.name == "Workflows")
        exceptions = tuple(Exception_(wf, status, category, end, free_) for wf, status, category, end, free_ in workflows.rows)
        config = next((dict((r[0], r[1]) for r in s.rows) for s in sheets if s.name == "Config"), {})
        starts = {kind: list_value(config.get(name)) for kind, name in START_PARAMETERS.items()}
        rules = effective_rules(("Team", "Feature"), catalog or {}, current or {}, set(free or ()), list(exceptions), starts)
        return cls({(r.workflow, r.status): StatusRule(r.category, bool(r.is_start), bool(r.is_end), bool(r.is_free)) for r in rules},
                   waits=has_wait(rules), entries=tuple(rules), exceptions=exceptions, start_statuses=starts,
                   known=frozenset(catalog or ()) | frozenset(current or ()))

    def listed(self) -> list:
        """The rules one by one (status_rules.Rule, source left empty): what status_rules.rule_counts reads."""
        from jira2dashboard.flow.status_rules import Rule
        return [Rule(wf, self._names[(wf, key)], r.category, int(r.is_start), int(r.is_end), int(r.is_free), "") for (wf, key), r in self._rows.items()]

    def has_start(self, workflow: str) -> bool:
        """A status of the workflow starts the flow (IsStart). Without one no ticket has a StartUtc: no Flow Time, Flow Efficiency or Flow Load."""
        return any(rule.is_start for (wf, _status), rule in self._rows.items() if wf == workflow)

    def ended_statuses(self) -> set[str]:
        """The statuses that end the flow (IsEnd) in at least one workflow: what the sheet Reconciliation lists with the tickets that sit in them."""
        return {self._names[k] for k, rule in self._rows.items() if rule.is_end}

    def rule(self, workflow: str, status: str) -> StatusRule:
        return self._rows.get((workflow, status_key(status)), StatusRule(UNCLASSIFIED, False, False, False))


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


def pi_sort_key(name: str) -> str:
    """The natural order of PI names (AL, 2026-11-25: any name is a PI, PIs without dates come in this order): a run of digits compares as a
    number (PI#3-2024 before PI#10-2026), a to z as A to Z. The same text as modFlowBuild.PiSortKey, compared as binary text; two names with
    the same key (PI#01-2027, PI#1-2027) go by the name itself."""
    out = []
    for n, part in enumerate(re.split(r"([0-9]+)", name)):
        if n % 2:
            digits = part.lstrip("0")
            out.append(f"\x01{len(digits):04d}{digits}")
        else:
            out.append("".join(chr(ord(c) - 32) if "a" <= c <= "z" else c for c in part))
    return "".join(out)


def ordered_pis(pis: dict) -> list[tuple]:
    """The PIs of the sheet PIs that have dates, (name, (start, end)), by Start, then by the natural order of their names (AL, 2026-11-25): the
    order of PIMetrics (modFlowBuild.PiNames), so also the first PI that holds a date (ItemFlow.DonePI) and the PIs of an iteration."""
    return sorted(pis.items(), key=lambda kv: (kv[1][0] is None, kv[1][0] or dt.datetime.min, pi_sort_key(kv[0]), kv[0]))


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
