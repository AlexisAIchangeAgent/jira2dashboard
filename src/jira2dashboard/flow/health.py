"""What the person should look at after a run: fields " | Name: value" of the status, each one only when there is something to see (a clean run
adds nothing). The Python twin of modHealth.HealthFields (same fields, same order, same words: tests/unit/test_health.py reads the VBA texts), for the
Python path (its sheet Run and python.log). They say WHY a figure or a chart is empty, with numbers and status or type names, never a value of a ticket:
  No start status                    a workflow kind that has tickets and no IsStart status: no StartUtc, so no Flow Time, Flow Efficiency or Flow Load
  No Wait status                     Flow Efficiency is not computable (Jira cannot say which in-progress statuses wait)
  Not Features                       tickets of the Program that are not Features (team "ART (other types)"), and a FeatureTypes that matches none
  Done outside the iterations        done tickets whose DoneUtc falls in no iteration (no metric counts them)
  PIs without dates                  PIs of the sheet PIs with no Start and no End: left out of the metrics (AL, 2026-11-25)
  Without a team                     tickets whose Team is empty
  Statuses not in Workflows          statuses of the history that no rule knows (Unclassified)
  Statuses unknown to Jira           a status that the sheet Workflows or Config FeatureStartStatuses / TeamStartStatuses names and no source of Jira
                                     knows (AL, 2026-11-27: the exception "À déployer" ruled nothing, Jira spells it "A déployer"); the WARN J2D-CFG-017
  Fields empty in every ticket       a ...Field of Config that is set while its column of Issues has no filled cell (AL, 2026-11-27: WorkTypeField =
                                     "Activité", a name, not the id of the field: Flow Distribution had nothing to distribute); the WARN J2D-CFG-018
  Started in the second of the backlog entry   no StartUtc because the start was entered in the same second as the last backlog entry (the
                                     milliseconds are dropped and the start must come strictly after it): an automation that moves a ticket at once
  AsOf before the data               status changes after AsOf: the open intervals and Flow Load stop at AsOf
  VelocityCheck gap                  rows of VelocityCheck whose Gap is not 0
  Reconciliation                     controls to review
The VBA also has Column types (raw_types: the sheet Run of the Python path has its own row) and Cells in error (formulas only).
AL, 2026-11-22: "flow metrics are still not computable".
"""
import datetime as dt
import re

from jira2dashboard import sheet_names as S
from jira2dashboard.flow import status_rules as SR
from jira2dashboard.flow.itemflow import ItemFlow
from jira2dashboard.flow.model import ART, BACKLOG, OTHER_TYPES, UNCLASSIFIED, WAIT, FlowIssue, Iteration, Rules
from jira2dashboard.flow.velocity_check import VelocityCheck

MAX_NAMES = 5                                   # names shown in a field, then ", ..."
HALF_SECOND = dt.timedelta(seconds=0.5)         # the margin of the formulas (refs.HALF)
KINDS = ("Feature", "Team")

# the literal parts of the texts, the same in modHealth.bas
NO_START = " (no Flow Time, Flow Efficiency or Flow Load): no status in progress, or a start status that Jira does not know in "
NO_WAIT = " | No Wait status (no Flow Efficiency): set Category = Wait on the statuses where the work waits in "
NOT_FEATURES = " tickets of the Program (team "
LEFT_OUT = ", left out of the ART metrics)"
MATCHES_NONE = """ matches no ticket of the Program (types read: """
TYPE_IT = "): type the name of the Feature type of Jira in "
OUTSIDE = " done tickets (no metric counts them): "
ADD_ITERATIONS = "add iterations to "
UNDATED = " | PIs without dates: "
UNDATED_WHY = "): left out of the metrics, type their Start and End in "
WITHOUT_TEAM = " tickets (a ticket takes its team from "
UNKNOWN = "): add them to "
SAME_SECOND = " | Started in the second of the backlog entry: "
SAME_SECOND_WHY = " tickets (no StartUtc, so left out of Flow Time and Flow Load)"
AS_OF = " | AsOf before the data: "
AS_OF_WHY = " status changes after AsOf (Flow Load stops at AsOf): set AsOf in "
AS_OF_FIX = " to the time of the extraction"
UNKNOWN_STATUSES = " | Statuses unknown to Jira: "
UNKNOWN_STATUSES_WHY = "): type each name as Jira spells it"
EMPTY_FIELDS = " | Fields empty in every ticket: "
NOT_AN_ID = ": not a field id"
EMPTY_FIELDS_WHY = "): type the id of the field (customfield_ and digits, listed in "
UNKNOWN_STATUSES_CODE, EMPTY_FIELDS_CODE = "J2D-CFG-017", "J2D-CFG-018"
# the ...Field parameters of Config and the column of Issues each one fills (modHealth.FIELD_COLUMNS: the same pairs)
FIELD_COLUMNS = (("WorkTypeField", "WorkType"), ("StoryPointsField", "StoryPoints"), ("PIField", "PI"), ("BusinessValueField", "BV"),
                 ("TeamField", "TeamValue"), ("SprintField", "SprintState"), ("TargetStartField", "TargetStart"), ("TargetEndField", "TargetEnd"))
_FIELD_ID = re.compile(r"[A-Za-z0-9_.]+")          # an id of Jira: customfield_10200, fixVersions...; a name has a space or an accent
GAP = " | VelocityCheck gap: "
REVIEW = " controls to review (sheet "
VBA_TEXTS = (" | No start status: ", NO_START, NO_WAIT, " | Not Features: ", NOT_FEATURES, LEFT_OUT, " | FeatureTypes ", TYPE_IT,
             " | Done outside the iterations: ", OUTSIDE, ADD_ITERATIONS, UNDATED, UNDATED_WHY, " | Without a team: ", WITHOUT_TEAM, " | Statuses not in Workflows: ", UNKNOWN,
             " with a Category", UNKNOWN_STATUSES, UNKNOWN_STATUSES_WHY, EMPTY_FIELDS, NOT_AN_ID, EMPTY_FIELDS_WHY, SAME_SECOND, SAME_SECOND_WHY, AS_OF,
             AS_OF_WHY, AS_OF_FIX, GAP, " rows (sheet ", " | Reconciliation: ", REVIEW)


def _names(names) -> tuple[str, int]:
    """The distinct names in their order of appearance, at most MAX_NAMES then ", ...", and how many there are."""
    distinct = list(dict.fromkeys(n for n in names if n))
    return ", ".join(distinct[:MAX_NAMES]) + (", ..." if len(distinct) > MAX_NAMES else ""), len(distinct)


def no_start_status(issues: list[FlowIssue], rules: Rules) -> str:
    with_tickets = {i.workflow for i in issues}
    starts = {r.workflow for r in rules.listed() if r.is_start}
    missing = [k for k in KINDS if k in with_tickets and k not in starts]
    return f" | No start status: {', '.join(missing)}{NO_START}{S.of('Config')}" if missing else ""


def no_wait_status(issues: list[FlowIssue], rules: Rules) -> str:
    if not issues or any(r.category == WAIT for r in rules.listed()):
        return ""
    return f"{NO_WAIT}{S.of('Workflows')}"


def not_features(items: list[ItemFlow], feature_types: str) -> str:
    other = [i for i in items if i.team == OTHER_TYPES]
    if not other:
        return ""
    art = sum(1 for i in items if i.team == ART)
    out = f" | Not Features: {len(other)} of {len(other) + art}{NOT_FEATURES}{OTHER_TYPES} in {S.of('ItemFlow')}{LEFT_OUT}"
    if art == 0:
        names, _n = _names(i.type.strip() for i in other)
        out += f' | FeatureTypes "{feature_types}"{MATCHES_NONE}{names}{TYPE_IT}{S.of("Config")}'
    return out


def outside_iterations(items: list[ItemFlow], iterations: list[Iteration]) -> str:
    done = [i for i in items if i.team != OTHER_TYPES and i.done is not None]
    out = sum(1 for i in done if not any(it.contains(i.done) for it in iterations))
    return f" | Done outside the iterations: {out} of {len(done)}{OUTSIDE}{ADD_ITERATIONS}{S.of('Iterations')}" if out else ""


def pis_without_dates(names) -> str:
    """The PIs of the sheet PIs with no Start and no End (CheckSettings.undated_pis), in the order of the sheet."""
    shown, n = _names(names)
    return f"{UNDATED}{n} ({shown}{UNDATED_WHY}{S.of('PIs')}" if n else ""


def without_team(items: list[ItemFlow]) -> str:
    n = sum(1 for i in items if not (i.team or "").strip())
    return f" | Without a team: {n}{WITHOUT_TEAM}{S.of('Projects')})" if n else ""


def unclassified_statuses(items: list[ItemFlow]) -> str:
    names, n = _names(iv.status for i in items for iv in i.intervals if iv.category == UNCLASSIFIED)
    return f" | Statuses not in Workflows: {n} ({names}{UNKNOWN}{S.of('Workflows')} with a Category" if n else ""


def unknown_statuses(rules: Rules, items: list[ItemFlow]) -> str:
    """The statuses that a setting names (the sheet Workflows, Config FeatureStartStatuses / TeamStartStatuses) and no source knows: the catalog, the
    current statuses, the histories. Nothing when Jira was not read at all (no catalog, no current status: nothing to compare with)."""
    if not rules.known:
        return ""
    known = set(rules.known) | {iv.status for i in items for iv in i.intervals}
    found = SR.unknown_statuses(rules.exceptions, rules.start_statuses, known)
    shown, n = _names(f"{name} ({S.of('Workflows') if where == 'Workflows' else where})" for name, where in found)
    return f"{UNKNOWN_STATUSES}{n} ({shown}{UNKNOWN_STATUSES_WHY}" if n else ""


def fields_filled(config: dict, header, rows) -> list[tuple[str, str, int]]:
    """(parameter, value, filled cells of its Issues column) for each ...Field of FIELD_COLUMNS, `config` the values of Config, `rows` the rows of
    Issues under `header`."""
    out = []
    for parameter, column in FIELD_COLUMNS:
        index = list(header).index(column)
        value = str(config.get(parameter) or "").strip()
        out.append((parameter, value, sum(1 for r in rows if r[index] not in (None, ""))))
    return out


def empty_fields(fields: list[tuple[str, str, int]], tickets: int) -> str:
    """The ...Field of Config that are set while no ticket holds a value (`fields`: fields_filled); "not a field id" when the value cannot be an id.
    Nothing when no ticket was read."""
    if tickets == 0:
        return ""
    empty = [(p, v) for p, v, filled in fields if v and filled == 0]
    shown, n = _names(p if _FIELD_ID.fullmatch(v) else f"{p} ({v}{NOT_AN_ID})" for p, v in empty)
    return f"{EMPTY_FIELDS}{n} ({shown}{EMPTY_FIELDS_WHY}{S.of('JiraFields')}) in {S.of('Config')}" if n else ""


def same_second_starts(issues: list[FlowIssue], items: list[ItemFlow], rules: Rules) -> str:
    """Tickets without a StartUtc that entered a start status in the same second as their last backlog entry."""
    n = 0
    for issue, item in zip(issues, items):
        backlog = [iv.entered for iv in item.intervals if iv.category == BACKLOG]
        if item.start is not None or not backlog:
            continue
        last = max(backlog)
        if any(rules.rule(issue.workflow, iv.status).is_start and abs(iv.entered - last) <= HALF_SECOND for iv in item.intervals):
            n += 1
    return f"{SAME_SECOND}{n}{SAME_SECOND_WHY}" if n else ""


def as_of_before_data(items: list[ItemFlow], as_of: dt.datetime) -> str:
    n = sum(1 for i in items for iv in i.intervals if iv.entered > as_of + HALF_SECOND)
    return f"{AS_OF}{n}{AS_OF_WHY}{S.of('Config')}{AS_OF_FIX}" if n else ""


def velocity_check_gap(checks: list[VelocityCheck]) -> str:
    n = sum(1 for c in checks if c.gap != 0)
    return f"{GAP}{n} of {len(checks)} rows (sheet {S.of('VelocityCheck')})" if n else ""


def reconciliation_findings(verdicts: list[str]) -> str:
    """The controls whose result is REVIEW (an INFO row is a point of attention: not counted)."""
    n = sum(1 for v in verdicts if v == "REVIEW")
    return f" | Reconciliation: {n} of {len(verdicts)}{REVIEW}{S.of('Reconciliation')})" if n else ""


def health_fields(issues: list[FlowIssue], items: list[ItemFlow], rules: Rules, iterations: list[Iteration], as_of: dt.datetime,
                  checks: list[VelocityCheck], verdicts: list[str], feature_types: str, undated_pis=(), fields=()) -> list[str]:
    """The fields in the order of modHealth.HealthFields; `items[n]` is the ItemFlow of `issues[n]`, `verdicts` the results of the Reconciliation,
    `feature_types` the text of Config.FeatureTypes, `undated_pis` the PIs of the sheet PIs without dates, `fields` the fields_filled of Config."""
    found = (no_start_status(issues, rules), no_wait_status(issues, rules), not_features(items, feature_types), outside_iterations(items, iterations),
             pis_without_dates(undated_pis), without_team(items), unclassified_statuses(items), unknown_statuses(rules, items),
             empty_fields(list(fields), len(issues)), same_second_starts(issues, items, rules), as_of_before_data(items, as_of),
             velocity_check_gap(checks), reconciliation_findings(verdicts))
    return [f for f in found if f]
