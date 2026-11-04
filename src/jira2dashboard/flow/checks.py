"""Data checks (docs/safe-flow-metrics-and-data-checks.md, section 3): every check except DC25 is a 0/1 flag per issue.

Pure. The Excel formulas (formulas.py: columns DC.. of ItemFlow, tables IssueLinks and FillRates, sheet DataChecks) are tested
against `run_checks` and `fill_rates`, the reference written here from the issue fields and the status history.

A check that needs a Jira field (Sprint, Team, PI, ...) is **not computable** while the id of that field is empty in
Config: its Result reads "not computable", never "OK" with 0 items, and no issue is flagged.

Interpretations (proposals; the document gives the idea, not always the exact rule):
* DC01 applies to the Team workflow, for the types of `Config.ParentRequiredTypes`. DC02 flags those issues when their
  parent is in the extract and its type is not in `Config.FeatureTypes`. DC03 flags the sub-tasks (`SubTaskTypes`) of a
  team whose sub-tasks have parents of more than one type. DC24 flags a feature (`FeatureTypes`) that no issue has as parent.
* DC04 and DC05 use the Jira status category (key `done`) and the resolution date.
* DC06 covers FROM_MISMATCH and CURRENT_MISMATCH (the history does not end in the current status). DC28 compares the
  status of two consecutive intervals with `Transitions` and `IsFree`.
* The sprint of an issue is the active one, else a future one, else the last closed one. DC16 counts the sprints in the
  Sprint field (3 or more = carried over twice), not the changelog. DC15: Jira category `indeterminate`, no sprint.
* DC17 and DC29 use the format `PI#` + 2 digits + `-` + 4 digits; for DC17 and DC18 a malformed value counts as empty.
  DC18 flags a feature when a child's sprint starts before the PI starts or ends after the PI ends (sheet `PIs`).
* DC19 applies to features whose current status category is not Backlog and not Removed (the document says "status >=
  Planifie"). DC20 flags an issue whose end is before its start, or whose dates are outside its parent's dates.
* DC22 flags the blocked issue when the sprint of its blocker ends after its own sprint starts. DC23 flags the issue
  without sprint on a link between two different teams.
* DC25: for every field x team x issue type with issues, the share of filled values; below `Config.FillRateThreshold` is
  one finding. DC27 flags issues whose key changed (a Key change in the changelog).
"""
import datetime as dt
import re
from collections import defaultdict
from dataclasses import dataclass, field

from jira2dashboard.flow.itemflow import item_flow
from jira2dashboard.flow.model import BACKLOG, REMOVED, UNCLASSIFIED, FlowIssue, Rules
from jira2dashboard.params.model import ParamSheet, load_all

PI_FORMAT = re.compile(r"PI#[0-9]{2}-[0-9]{4}")
PIS_SHEET = "PIs"                                    # a need that is a sheet with rows, not a Config parameter


@dataclass(frozen=True)
class Check:
    id: str
    title: str
    severity: str
    needs: tuple[str, ...] = ()                      # Config parameters (ids of fields) or PIS_SHEET
    source: str = "ItemFlow"                         # where the items are counted: the column of ItemFlow, or FillRates


CHECKS = (
    Check("DC01", "Story or task without parent", "Blocker"),
    Check("DC02", "Parent is not a feature (wrong level)", "Major"),
    Check("DC03", "Sub-tasks of a team have parents of different types", "Minor"),
    Check("DC04", "Done without resolution date", "Major"),
    Check("DC05", "Resolution date but not done", "Major"),
    Check("DC06", "Status history inconsistent (FROM_MISMATCH or CURRENT_MISMATCH)", "Major"),
    Check("DC07", "Status change dated before the creation of the issue (BEFORE_CREATED)", "Major"),
    Check("DC08", "Status not in Workflows (category Unclassified)", "Blocker for efficiency"),
    Check("DC09", "Done without ever being in a start status", "Major"),
    Check("DC10", "Reopened after Done", "Major"),
    Check("DC11", "Story without story points in an active or future sprint", "Blocker for capacity",
          ("StoryPointsField", "SprintField")),
    Check("DC12", "Story points on sub-tasks", "Major", ("StoryPointsField",)),
    Check("DC13", "No Team value", "Blocker for Plans", ("TeamField",)),
    Check("DC14", "Team value does not match the team project", "Major", ("TeamField",)),
    Check("DC15", "In progress without sprint", "Major", ("SprintField",)),
    Check("DC16", "Carried over 2+ sprints", "Minor (signal)", ("SprintField",)),
    Check("DC17", "Feature without PI value", "Blocker for PI views", ("PIField",)),
    Check("DC18", "Feature PI inconsistent with its stories' sprints", "Major", ("PIField", "SprintField", PIS_SHEET)),
    Check("DC19", "Feature without fixed dates (Target start / Target end)", "Blocker for Plans timeline",
          ("TargetStartField", "TargetEndField")),
    Check("DC20", "Target end before Target start, or child dates outside parent dates", "Major",
          ("TargetStartField", "TargetEndField")),
    Check("DC21", "Business value on a feature", "Major", ("BusinessValueField",)),
    Check("DC22", "Dependency risk: blocker ends after the blocked item starts", "Major", ("SprintField",)),
    Check("DC23", "Cross-team dependency without both items planned", "Minor", ("SprintField",)),
    Check("DC24", "Feature without children", "Minor"),
    Check("DC25", "Field fill rate below threshold", "Major", (), "FillRates"),
    Check("DC26", "Missing fix version on features", "Minor"),
    Check("DC27", "Issue moved between projects (key changed)", "Minor"),
    Check("DC28", "Status transition not in Transitions", "Major"),
    Check("DC29", "PI field not empty but not in the format PI#XX-YYYY", "Major", ("PIField",)),
)
IDS = tuple(c.id for c in CHECKS)
PER_ISSUE = tuple(c.id for c in CHECKS if c.source == "ItemFlow")      # DC25 is counted on the table FillRates
FIELD_PARAMETERS = ("StoryPointsField", "BusinessValueField", "PIField", "TeamField", "SprintField", "TargetStartField",
                    "TargetEndField")

# The fields of the fill-rate matrix (DC25): ItemFlow column -> Config parameter that must be set ("" = a system field)
FILL_FIELDS = (("Parent", ""), ("FixVersions", ""), ("ResolvedUtc", ""), ("StoryPoints", "StoryPointsField"),
               ("PI", "PIField"), ("TeamValue", "TeamField"), ("SprintState", "SprintField"),
               ("TargetStart", "TargetStartField"), ("TargetEnd", "TargetEndField"), ("BV", "BusinessValueField"))


def _list(value) -> frozenset[str]:
    return frozenset(t.strip() for t in str(value or "").split(",") if t.strip())


@dataclass(frozen=True)
class CheckSettings:
    """What the checks read from the workbook: the Config rows and the sheet PIs."""
    parent_types: frozenset[str] = frozenset()
    feature_types: frozenset[str] = frozenset()
    subtask_types: frozenset[str] = frozenset()
    points_types: frozenset[str] = frozenset()
    fields: dict[str, str] = field(default_factory=dict)         # Config parameter -> field id ("" = not configured)
    fill_threshold: int = 90
    pis: dict[str, tuple[dt.datetime, dt.datetime]] = field(default_factory=dict)

    def configured(self, name: str) -> bool:
        return bool(self.fields.get(name))

    def ready(self, check: Check) -> bool:
        """False while a field that the check reads is not configured (the check is "not computable")."""
        return all(bool(self.pis) if need == PIS_SHEET else self.configured(need) for need in check.needs)

    @classmethod
    def from_params(cls, sheets: list[ParamSheet] | None = None) -> "CheckSettings":
        by_name = {s.name: s for s in (sheets or load_all())}
        config = {row[0]: row[1] for row in by_name["Config"].rows}
        pis = {name: (dt.datetime.fromisoformat(start), dt.datetime.fromisoformat(end))
               for name, start, end in by_name["PIs"].rows}
        return cls(parent_types=_list(config["ParentRequiredTypes"]), feature_types=_list(config["FeatureTypes"]),
                   subtask_types=_list(config["SubTaskTypes"]), points_types=_list(config["PointsRequiredTypes"]),
                   fields={name: str(config.get(name) or "") for name in FIELD_PARAMETERS},
                   fill_threshold=int(config["FillRateThreshold"]), pis=pis)


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


def valid_pi(value: str | None) -> bool:
    return bool(value) and PI_FORMAT.fullmatch(value) is not None


def _status_history(issue: FlowIssue, as_of: dt.datetime, rules: Rules, moves: Moves) -> set[str]:
    """DC06 to DC10 and DC28: from the status changes."""
    found: set[str] = set()
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
    return found


def run_checks(issues: list[FlowIssue], as_of: dt.datetime, rules: Rules, moves: Moves,
               settings: CheckSettings) -> dict[str, frozenset[str]]:
    """{issue key: ids of the checks it fails}. A check that is not ready (see CheckSettings.ready) flags nobody."""
    s = settings
    by_key = {i.key: i for i in issues}
    children: dict[str, list[FlowIssue]] = defaultdict(list)
    for item in issues:
        if item.parent:
            children[item.parent].append(item)
    subtask_parent_types: dict[str, set[str]] = defaultdict(set)         # team -> types of the parents of its sub-tasks
    for item in issues:
        parent = by_key.get(item.parent or "")
        if item.type in s.subtask_types and parent is not None:
            subtask_parent_types[item.team].add(parent.type)
    on_links = _link_flags(issues, by_key)
    ready = {c.id: s.ready(c) for c in CHECKS}
    flagged: dict[str, frozenset[str]] = {}
    for issue in issues:
        found = _status_history(issue, as_of, rules, moves)
        found |= _field_checks(issue, by_key.get(issue.parent or ""), children[issue.key], subtask_parent_types, rules, s)
        found |= on_links.get(issue.key, set())
        flagged[issue.key] = frozenset(i for i in found if ready[i])
    return flagged


def _field_checks(issue: FlowIssue, parent: FlowIssue | None, children: list[FlowIssue],
                  subtask_parent_types: dict[str, set[str]], rules: Rules, s: CheckSettings) -> set[str]:
    found: set[str] = set()
    done = issue.status_category == "done"
    feature = issue.type in s.feature_types
    team_issue = issue.workflow == "Team"
    if team_issue and issue.type in s.parent_types and not issue.parent:
        found.add("DC01")
    if team_issue and issue.type in s.parent_types and parent is not None and parent.type not in s.feature_types:
        found.add("DC02")
    if issue.type in s.subtask_types and parent is not None and len(subtask_parent_types[issue.team]) > 1:
        found.add("DC03")
    if done and issue.resolved is None:
        found.add("DC04")
    if not done and issue.resolved is not None:
        found.add("DC05")
    if issue.type in s.points_types and issue.sprint_state in ("active", "future") and issue.story_points is None:
        found.add("DC11")
    if issue.type in s.subtask_types and issue.story_points is not None:
        found.add("DC12")
    if (feature or issue.type in s.parent_types) and not issue.team_value:
        found.add("DC13")
    if team_issue and issue.team_value and issue.team_value.casefold() != issue.team.casefold():
        found.add("DC14")
    if team_issue and issue.status_category == "indeterminate" and issue.sprint_count == 0:
        found.add("DC15")
    if issue.sprint_count >= 3:
        found.add("DC16")
    if feature and not valid_pi(issue.pi):
        found.add("DC17")
    if feature and valid_pi(issue.pi) and issue.pi in s.pis:
        start, end = s.pis[issue.pi]
        if any(c.sprint_start is not None and (c.sprint_start < start or c.sprint_end > end) for c in children):
            found.add("DC18")
    category = rules.rule(issue.workflow, issue.status).category
    if feature and category not in (BACKLOG, REMOVED) and (issue.target_start is None or issue.target_end is None):
        found.add("DC19")
    if _dates_wrong(issue, parent):
        found.add("DC20")
    if feature and issue.business_value is not None:
        found.add("DC21")
    if feature and not children:
        found.add("DC24")
    if feature and not issue.fix_versions:
        found.add("DC26")
    if issue.key_changes > 0:
        found.add("DC27")
    if issue.pi and not valid_pi(issue.pi):
        found.add("DC29")
    return found


def _dates_wrong(issue: FlowIssue, parent: FlowIssue | None) -> bool:
    start, end = issue.target_start, issue.target_end
    if start is not None and end is not None and end < start:
        return True
    if parent is None:
        return False
    if start is not None and parent.target_start is not None and start < parent.target_start:
        return True
    return end is not None and parent.target_end is not None and end > parent.target_end


def blocking_links(issues: list[FlowIssue]) -> list[tuple[FlowIssue, FlowIssue]]:
    """(blocker, blocked) for every link of type Blocks between two issues of the extract."""
    by_key = {i.key: i for i in issues}
    return [(i, by_key[k]) for i in issues for k in i.blocks if k in by_key]


def _link_flags(issues: list[FlowIssue], by_key: dict[str, FlowIssue]) -> dict[str, set[str]]:
    """DC22 on the blocked issue; DC23 on the side(s) without sprint of a link between two teams."""
    found: dict[str, set[str]] = defaultdict(set)
    for blocker, blocked in blocking_links(issues):
        if blocker.sprint_end is not None and blocked.sprint_start is not None and blocker.sprint_end > blocked.sprint_start:
            found[blocked.key].add("DC22")
        if blocker.team != blocked.team:
            for side in (blocker, blocked):
                if side.sprint_count == 0:
                    found[side.key].add("DC23")
    return found


@dataclass(frozen=True)
class FillRate:
    field: str
    team: str
    type: str
    count: int
    filled: int
    threshold: int

    @property
    def below(self) -> bool:
        return self.count > 0 and self.filled * 100 < self.count * self.threshold


def _filled(issue: FlowIssue, name: str) -> bool:
    value = {"Parent": issue.parent, "FixVersions": issue.fix_versions, "ResolvedUtc": issue.resolved,
             "StoryPoints": issue.story_points, "PI": issue.pi, "TeamValue": issue.team_value,
             "SprintState": issue.sprint_state, "TargetStart": issue.target_start, "TargetEnd": issue.target_end,
             "BV": issue.business_value}[name]
    return value not in (None, "", ())


def fill_rates(issues: list[FlowIssue], settings: CheckSettings, teams: list[str], types: list[str]) -> list[FillRate]:
    """DC25: one row per field (a system field, or one that is configured) x team x issue type, in that order."""
    out = []
    for name, need in FILL_FIELDS:
        if need and not settings.configured(need):
            continue
        for team in teams:
            for type_ in types:
                group = [i for i in issues if i.team == team and i.type == type_]
                out.append(FillRate(name, team, type_, len(group), sum(1 for i in group if _filled(i, name)),
                                    settings.fill_threshold))
    return out
