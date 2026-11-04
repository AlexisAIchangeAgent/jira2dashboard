"""Deterministic volume data for the fake Jira: a Program project (features) and team projects (stories, tasks,
bugs, sub-tasks), with status histories that follow the transitions of params/transitions.md.

Same seed, same dataset. The histories include the cases the flow metrics must handle: a step back, a detour
through Dépriorisé then Nouveau (the flow time restarts), a cancellation, issues still open.
"""
import datetime as dt
import random
from collections import defaultdict

from jira2dashboard.mockjira.dataset import Change, Dataset, FieldIds, Issue, StatusCatalog
from jira2dashboard.params.model import ParamSheet, load_all

DEFAULT_PROJECTS = ("PRG", "TA", "TB")      # the first project is the Program (features), the others are teams
AUTHORS = (("Ann Martin", "acc-ann"), ("Bob Durand", "acc-bob"), ("Chloe Petit", "acc-chloe"),
           ("Dan Roux", "acc-dan"), ("Eva Blanc", "acc-eva"))
P_STEP_BACK, P_DETOUR, P_CANCEL_AT_DETOUR = 0.12, 0.06, 0.3


class _Workflow:
    """The moves of one workflow, read from the parameter sheets."""

    def __init__(self, name: str, sheets: dict[str, ParamSheet]):
        self.forward: dict[str, str] = {}
        self.back: dict[str, list[str]] = defaultdict(list)
        self.exit_to: str | None = None
        self.restart_at: str | None = None
        for wf, source, target, kind in sheets["Transitions"].rows:
            if wf != name:
                continue
            if kind == "forward":
                self.forward[source] = target
            elif kind == "back":
                self.back[source].append(target)
            elif kind == "exit":
                self.exit_to = target
            elif kind == "back to backlog":
                self.restart_at = target
        rows = [r for r in sheets["Workflows"].rows if r[0] == name]
        self.start = next(r[2] for r in rows if r[4] == 1)
        self.end = next(r[2] for r in rows if r[5] == 1)
        self.free = next(r[2] for r in rows if r[6] == 1)
        self.first = next(r[2] for r in rows if r[1] == 1)
        self.chain = [r[2] for r in sorted((r for r in rows if r[1] is not None), key=lambda r: r[1])]


def _walk(rng: random.Random, wf: _Workflow, created: dt.datetime, as_of: dt.datetime) -> list[tuple[dt.datetime, str]]:
    """The moves (time, status entered) of one issue, never after `as_of`."""
    stop_at = rng.choice(wf.chain[: wf.chain.index(wf.end) + 1] + [wf.end] * 2)
    status, now, moves = wf.first, created, []

    def go(target: str) -> bool:
        nonlocal status, now
        now = now + dt.timedelta(days=rng.uniform(0.2, 9.0))
        if now > as_of:
            return False
        moves.append((now.replace(microsecond=0), target))
        status = target
        return True

    while status != stop_at:
        roll = rng.random()
        if roll < P_STEP_BACK and wf.back.get(status):
            if not go(rng.choice(wf.back[status])):
                break
        elif roll < P_STEP_BACK + P_DETOUR:                     # Dépriorisé is reachable from any status
            if not go(wf.free):
                break
            if rng.random() < P_CANCEL_AT_DETOUR:
                go(wf.exit_to)
                break
            if not go(wf.restart_at):                           # back to Nouveau: the flow time restarts
                break
        elif not go(wf.forward[status]):
            break
    return moves


def _new_issue(rng, key, issue_id, project, type_, work_type, workflow_name, wf, created, as_of, parent):
    moves = _walk(rng, wf, created, as_of)
    changes, previous = [], wf.first
    for n, (at, status) in enumerate(moves, start=1):
        who, who_id = rng.choice(AUTHORS)
        changes.append(Change(id=f"{issue_id}{n:03d}", at=at, from_status=previous, to_status=status,
                              author=who, author_id=who_id))
        previous = status
    return Issue(
        key=key, id=str(issue_id), project=project, type=type_, work_type=work_type,
        summary=f"{type_} {key}", created=created.replace(microsecond=0), changes=tuple(changes),
        initial_status=wf.first, workflow=workflow_name, parent=parent,
        priority=rng.choice(("Low", "Medium", "High")), assignee=rng.choice(AUTHORS)[0],
        labels=tuple(rng.sample(("api", "ui", "data", "ops"), rng.randint(0, 2))),
        story_points=rng.choice((1, 2, 3, 5, 8)) if type_ in ("Story", "Task") else None,
        business_value=rng.randint(1, 100) if type_ == "Feature" else None,
        pi=rng.choice(("PI#01-2026", "PI#02-2026", None)),
        due=None)


def generate(seed: int = 1, issues: int = 1200, as_of: dt.datetime = dt.datetime(2026, 11, 3, 18, 0),
             since: dt.datetime = dt.datetime(2026, 6, 1), sheets: list[ParamSheet] | None = None,
             projects: tuple[str, ...] = DEFAULT_PROJECTS, fields: FieldIds | None = None) -> Dataset:
    """`issues` issues: about 10 % features in the first project (Program, Feature workflow), the rest in the
    other projects (teams, Team workflow). `fields` sets the ids of the custom fields (story points, ...)."""
    if len(projects) < 2:
        raise ValueError("need a Program project and at least one team project")
    program, teams = projects[0], projects[1:]
    rng = random.Random(seed)
    sheets_by_name = {s.name: s for s in (sheets or load_all())}
    workflows = {name: _Workflow(name, sheets_by_name) for name in ("Team", "Feature")}
    span = (as_of - since - dt.timedelta(days=1)).total_seconds()
    created_at = sorted(since + dt.timedelta(seconds=rng.uniform(0, span)) for _ in range(issues))
    counters: dict[str, int] = defaultdict(int)
    features: list[str] = []
    stories: dict[str, list[str]] = defaultdict(list)
    result: list[Issue] = []
    for n, created in enumerate(created_at, start=1):
        roll = rng.random()
        if roll < 0.10 or n == 1:
            project, type_, work_type, wf_name = program, "Feature", rng.choice(("Feature", "Enabler")), "Feature"
        else:
            project = rng.choice(teams)
            type_, work_type = rng.choices(
                (("Story", "Feature"), ("Story", "Enabler"), ("Task", "Autre"), ("Bug", "Defect"), ("Sub-task", "Autre")),
                weights=(40, 15, 20, 10, 15))[0]
            wf_name = "Team"
        if type_ == "Sub-task" and not stories[project]:
            type_, work_type = "Task", "Autre"
        counters[project] += 1
        key = f"{project}-{counters[project]}"
        parent = None
        if type_ == "Story" and features and rng.random() < 0.85:
            parent = rng.choice(features)
        elif type_ == "Sub-task":
            parent = rng.choice(stories[project])
        item = _new_issue(rng, key, 10000 + n, project, type_, work_type, wf_name, workflows[wf_name], created,
                          as_of, parent)
        result.append(item)
        if type_ == "Feature":
            features.append(key)
        elif type_ == "Story":
            stories[project].append(key)
    return Dataset(tuple(result), as_of, StatusCatalog.from_params(list(sheets_by_name.values())),
                   fields or FieldIds())
