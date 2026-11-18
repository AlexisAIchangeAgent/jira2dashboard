"""The scenario art-metrics: 25 Features of the Program project over 4 PIs, with a known history (AL, 2026-11-16).

It exists so that every flow metric and chart of the workbook can be checked against numbers worked out by hand from this
text (tests/unit/test_art_metrics.py types them; nothing is computed from flow/ there):
- PIs are consecutive: 18 weeks each from 2025-01-01, cut in 6 iterations of 3 weeks (IT-<pi>-<n>), Start included, End
  excluded, no gap. PI-1, PI-2, PI-3 and PI-4 hold 6 Features each at their start (ft-1..6, 7..12, 13..18, 19..24).
- Workflow program-workflow (named so in Jira; the sheets keep the workflow value Feature, the only one the parameters know):
  todo, analyzing, ready4PI, doing, ready2test, testing, ready2deploy, deploying, released, and the free status deferred.
- Every status lasts 1 day except one per PI: testing (PI-1), analyzing (PI-2), doing (PI-3), ready2deploy (PI-4): 2 days.
  A Feature is released (or deferred) at noon of day 10 of its iteration; the chain of statuses is laid back from there.
- A deferred Feature leaves its PI in the same instant: its PI field is emptied or moved to PI-5 (a PI the sheet PIs lacks).
  The PI field of PI-2, PI-3 and PI-4 Features is set the day before their PI starts (history); the Iteration field stays.
- ft-25 is created during IT-4-2 with PI-4 and IT-4-4 as first values and released in IT-4-6: added after the start of its PI.
- 3 teams in the Team field: ft-1,2 team-1; ft-3,4 team-2; ft-5,6 team-3, then team-1, 2, 3 in turn within each PI; ft-25 team-1.
- 3 sub-tasks of the Program (TTT-26, 27 under ft-7, ft-8, done in IT-2-1; TTT-28 under ft-16, done in IT-3-4): not Features, so the ART leaves them out
  (team "ART (other types)", AL 2026-11-18); every figure of the 25 Features is unchanged.
"""
import datetime as dt
from dataclasses import dataclass

from jira2dashboard.flow.model import Iteration
from jira2dashboard.mockjira.dataset import Dataset, FieldChange, FieldIds, ParamSheet, Scenario, StatusCatalog, issue
from jira2dashboard.params.model import load_all

NAME = "art-metrics"
PORT = 8101                                             # its own fake Jira instance: the generic one (8099) can run beside it
START = dt.datetime(2025, 1, 1)
PI_WEEKS, ITERATION_WEEKS, PI_COUNT = 18, 3, 4
PROGRAM = "TTT"                                         # the one project: the Program that holds the Features
# only the fields the scenario uses: no story points, business value, sprint or target dates (not in its catalog, not in its issues)
FIELDS = FieldIds(story_points=None, business_value=None, sprint=None, target_start=None, target_end=None,
                  iteration="customfield_10500")
JIRA_WORKFLOW = "program-workflow"
WORKFLOW = "Feature"                                    # the workflow value of the parameter sheets (params/model.py)
TODO, DONE, DEFERRED = "todo", "released", "deferred"
CHAIN = ("analyzing", "ready4PI", "doing", "ready2test", "testing", "ready2deploy", "deploying")
LONG_STATUS = {1: "testing", 2: "analyzing", 3: "doing", 4: "ready2deploy"}          # per PI number: the status that lasts 2 days
RELEASE_DAY = dt.timedelta(days=10, hours=12)                                         # in its iteration
SETTLED = "2025-01-01T00:00"                                                          # the backlog, created at the start of PI-1
ADDED = "2026-02-10T09:00"                                                            # ft-25, during IT-4-2

# status, Workflows rows (Workflow, Status, Category, IsStart, IsEnd, IsFree) and Transitions (Workflow, From, To)
_STATUS_ROWS = ((TODO, "Backlog", 0, 0, 0), ("analyzing", "Active", 1, 0, 0), ("ready4PI", "Wait", 0, 0, 0),
                ("doing", "Active", 0, 0, 0), ("ready2test", "Wait", 0, 0, 0), ("testing", "Active", 0, 0, 0),
                ("ready2deploy", "Wait", 0, 0, 0), ("deploying", "Active", 0, 0, 0), (DONE, "Done", 0, 1, 0))
WORKFLOW_ROWS = tuple((WORKFLOW, name, category, start, end, free) for name, category, start, end, free in _STATUS_ROWS) + (
    (WORKFLOW, DEFERRED, "Backlog", 0, 0, 1),)
TRANSITION_ROWS = tuple((WORKFLOW, _STATUS_ROWS[n][0], _STATUS_ROWS[n + 1][0]) for n in range(len(_STATUS_ROWS) - 1)) + (
    (WORKFLOW, DEFERRED, TODO),)


@dataclass(frozen=True)
class Plan:
    number: int                         # ft-<number>
    pi: int                             # the PI it is engaged in
    team: str
    iteration: int                      # 1..6: the iteration of the PI where it is released or deferred
    outcome: str                        # DONE or DEFERRED
    moved_to: str | None = None         # the PI field after a deferral (None = emptied)
    added: bool = False                 # created during the PI, with its PI and iteration as first values


PLANS = (
    Plan(1, 1, "team-1", 1, DONE), Plan(2, 1, "team-1", 2, DONE), Plan(3, 1, "team-2", 3, DONE),
    Plan(4, 1, "team-2", 4, DONE), Plan(5, 1, "team-3", 5, DONE), Plan(6, 1, "team-3", 6, DONE),
    Plan(7, 2, "team-1", 1, DONE), Plan(8, 2, "team-2", 2, DONE), Plan(9, 2, "team-3", 3, DONE),
    Plan(10, 2, "team-1", 4, DONE), Plan(11, 2, "team-2", 5, DONE), Plan(12, 2, "team-3", 6, DEFERRED, None),
    Plan(13, 3, "team-1", 1, DEFERRED, "PI-5"), Plan(14, 3, "team-2", 2, DEFERRED, None), Plan(15, 3, "team-3", 3, DEFERRED, "PI-5"),
    Plan(16, 3, "team-1", 4, DONE), Plan(17, 3, "team-2", 4, DONE), Plan(18, 3, "team-3", 6, DONE),
    Plan(19, 4, "team-1", 1, DONE), Plan(20, 4, "team-2", 2, DONE), Plan(21, 4, "team-3", 3, DONE),
    Plan(22, 4, "team-1", 5, DONE), Plan(23, 4, "team-2", 4, DEFERRED, None), Plan(24, 4, "team-3", 6, DEFERRED, "PI-5"),
    Plan(25, 4, "team-1", 6, DONE, added=True),
)


def pis() -> tuple[Iteration, ...]:
    """PI-1.. one after the other, 18 weeks each (End excluded: the next one starts there)."""
    start, out = START, []
    for n in range(1, PI_COUNT + 1):
        end = start + dt.timedelta(weeks=PI_WEEKS)
        out.append(Iteration(f"PI-{n}", start, end))
        start = end
    return tuple(out)


def iterations() -> tuple[Iteration, ...]:
    """IT-<pi>-<n>: 6 iterations of 3 weeks in each PI."""
    out = []
    for n, pi in enumerate(pis(), start=1):
        for k in range(6):
            start = pi.start + dt.timedelta(weeks=ITERATION_WEEKS * k)
            out.append(Iteration(f"IT-{n}-{k + 1}", start, start + dt.timedelta(weeks=ITERATION_WEEKS)))
    return tuple(out)


def sheets() -> tuple[ParamSheet, ...]:
    """The parameter sheets of the scenario, in the order of the workbook. Projects, Iterations, PIs, Workflows and Transitions are
    its own (built here: the Program project only, its calendar, the Feature workflow only: no Team project, no Team workflow, nothing
    of params/*.md). Config is the product's own list of parameters: it comes from params/."""
    own = {"Projects": ((PROGRAM, "Program", None),),
           "Iterations": tuple((i.name, f"{i.start:%Y-%m-%d}", f"{i.end:%Y-%m-%d}") for i in iterations()),
           "PIs": tuple((p.name, f"{p.start:%Y-%m-%d}", f"{p.end:%Y-%m-%d}") for p in pis()),
           "Workflows": WORKFLOW_ROWS, "Transitions": TRANSITION_ROWS}
    return tuple(ParamSheet(sheet.name, sheet.header, own.get(sheet.name, sheet.rows)) for sheet in load_all())


def _stamp(moment: dt.datetime) -> str:
    return moment.isoformat(timespec="minutes")


def _moves(plan: Plan, released_at: dt.datetime) -> tuple[tuple[str, str], ...]:
    """(time, status) of each move: the chain laid back from the release (or the deferral, after doing)."""
    chain = CHAIN if plan.outcome == DONE else CHAIN[:CHAIN.index("doing") + 1]
    long_status = LONG_STATUS[plan.pi]
    entered, at = [], released_at
    for status in reversed(chain):
        at -= dt.timedelta(days=2 if status == long_status else 1)
        entered.append((_stamp(at), status))
    return tuple(reversed(entered)) + ((_stamp(released_at), plan.outcome),)


def _feature(plan: Plan, project: str, all_iterations: tuple[Iteration, ...], all_pis: tuple[Iteration, ...], fields: FieldIds):
    label = f"IT-{plan.pi}-{plan.iteration}"
    found = next(i for i in all_iterations if i.name == label)
    at = found.start + RELEASE_DAY                                                    # released or deferred at this moment
    pi = all_pis[plan.pi - 1]
    name = f"PI-{plan.pi}"
    changes, first_pi = [], (name if plan.pi == 1 or plan.added else None)
    if first_pi is None:                                                               # the PI is set the day before it starts
        planned_at = pi.start - dt.timedelta(days=1) + dt.timedelta(hours=9)
        changes.append(FieldChange(f"{plan.number}f1", planned_at, fields.pi, "PI", None, name))
    if plan.outcome == DEFERRED:
        changes.append(FieldChange(f"{plan.number}f2", at, fields.pi, "PI", name, plan.moved_to))
    current = plan.moved_to if plan.outcome == DEFERRED else name
    return issue(f"{project}-{plan.number}", moves=_moves(plan, at), created=ADDED if plan.added else SETTLED,
                 initial_status=TODO, type="Feature", work_type="Feature", summary=f"ft-{plan.number}", workflow=WORKFLOW,
                 pi=current, team=plan.team, iteration=label if not plan.added else "IT-4-4", field_changes=tuple(changes))


# Sub-tasks of the Program project (TTT-26.. under Features): tickets that are NOT Features. Config.FeatureTypes = Feature, so the team of the ART does not
# count them ("ART (other types)"): none of the figures above changes. (key number, parent Feature number, iteration where it is done at 10 d 12 h)
SUBTASKS = ((26, 7, "IT-2-1"), (27, 8, "IT-2-1"), (28, 16, "IT-3-4"))


def _subtask(number: int, parent: int, label: str, project: str, all_iterations: tuple[Iteration, ...]):
    at = next(i for i in all_iterations if i.name == label).start + RELEASE_DAY
    moves = ((_stamp(at - dt.timedelta(days=2)), "analyzing"), (_stamp(at - dt.timedelta(days=1)), "doing"), (_stamp(at), DONE))
    return issue(f"{project}-{number}", moves=moves, created=SETTLED, initial_status=TODO, type="Sub-task", work_type="Autre",
                 summary=f"sub-task of ft-{parent}", workflow=WORKFLOW, parent=f"{project}-{parent}")


def scenario() -> Scenario:
    project = PROGRAM
    all_pis, all_iterations = pis(), iterations()
    own = sheets()
    issues = tuple(_feature(plan, project, all_iterations, all_pis, FIELDS) for plan in PLANS) + tuple(
        _subtask(n, parent, label, project, all_iterations) for n, parent, label in SUBTASKS)
    dataset = Dataset(issues, all_pis[-1].end, catalog=StatusCatalog.from_params(list(own)), fields=FIELDS, sheets=own,
                      workflow_names={WORKFLOW: JIRA_WORKFLOW})
    return Scenario(NAME, dataset, all_pis, all_iterations, port=PORT)
