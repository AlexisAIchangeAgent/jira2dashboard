"""Issues with known data anomalies, to test the data checks (flow/checks.py) on the fake Jira.

Each case is written by hand with the ids of the checks it must trigger. A case starts from a fully clean issue (a parent,
a team, a sprint, story points...) and breaks one thing, so that its expected ids are short and certain. Features come
with one clean child each (otherwise they would fail DC24). The generated issues are not clean on purpose (they have gaps):
a test compares the cases with their expectation and every other issue with the Python reference.
"""
import dataclasses
import datetime as dt

from jira2dashboard.mockjira.dataset import Dataset, Issue, Sprint, issue
from jira2dashboard.params.model import load_all

FIRST_NUMBER = 9001                       # keys of the cases: <project>-9001 ... (the generated ones stay far below)
CHAIN = (("2026-10-02T10:00", "En analyse"), ("2026-10-03T10:00", "Prêt"), ("2026-10-04T10:00", "Planifié"),
         ("2026-10-05T10:00", "En développement"), ("2026-10-06T10:00", "En test"),
         ("2026-10-07T10:00", "À déployer"), ("2026-10-08T10:00", "Terminé"))
ACTIVE = Sprint(1, "Sprint 1", "active", dt.datetime(2026, 10, 26), dt.datetime(2026, 11, 9))
CLOSED = (Sprint(2, "Sprint 2", "closed", dt.datetime(2026, 9, 14), dt.datetime(2026, 9, 28)),
          Sprint(3, "Sprint 3", "closed", dt.datetime(2026, 9, 28), dt.datetime(2026, 10, 12)))
BEFORE_PI = Sprint(9, "Sprint 9", "closed", dt.datetime(2026, 5, 18), dt.datetime(2026, 6, 2))
PI = "PI#01-2026"
PI_WINDOW = (dt.datetime(2026, 6, 1), dt.datetime(2026, 12, 31))      # the row of the sheet PIs: it holds every sprint above but one
FEATURE_DATES = dict(target_start=dt.date(2026, 9, 1), target_end=dt.date(2026, 12, 20))


def team_of(project: str) -> str:
    projects = next(s for s in load_all() if s.name == "Projects")
    return next(team for key, _role, team in projects.rows if key == project)


def known_anomalies(projects: tuple[str, ...]) -> dict[Issue, frozenset[str]]:
    """The cases, in the Program project (features) and two team projects, with the ids of the checks they fail."""
    if len(projects) < 3:
        raise ValueError("the known anomalies need the Program and two team projects (params/projects.md has them)")
    program, alpha, beta = projects[0], projects[1], projects[2]
    cases: list[tuple[Issue, frozenset[str]]] = []
    counter = [FIRST_NUMBER - 1]

    def key(project: str) -> str:
        counter[0] += 1
        return f"{project}-{counter[0]}"

    def add(item: Issue, ids: tuple[str, ...]) -> Issue:
        cases.append((item, frozenset(ids)))
        return item

    def feature(*ids: str, **fields) -> Issue:
        base = dict(type="Feature", moves=CHAIN[:2], work_type="Feature", pi=PI, team="ART", fix_versions=("R1",),
                    **FEATURE_DATES)
        return add(issue(key(program), **{**base, **fields}), ids)

    def story(parent: str | None, *ids: str, project: str = alpha, **fields) -> Issue:
        base = dict(moves=CHAIN[:2], parent=parent, story_points=3, team=team_of(project), sprints=(ACTIVE,))
        return add(issue(key(project), project=project, **{**base, **fields}), ids)

    clean = feature()                                               # a clean feature: the parent of the clean stories
    story(clean.key)                                                # a clean control
    story(clean.key, "DC09", "DC28", moves=(("2026-10-02T10:00", "Terminé"),))
    story(clean.key, "DC10", "DC28", moves=CHAIN + (("2026-10-09T10:00", "En développement"),))
    story(clean.key, "DC08", "DC28", moves=CHAIN[:1] + (("2026-10-03T10:00", "Statut inconnu"),))
    story(clean.key, "DC07", moves=(("2026-09-30T10:00", "En analyse"), ("2026-10-02T10:00", "Prêt")))
    story(clean.key, "DC04", moves=CHAIN, resolution_override=(None,))
    story(clean.key, "DC05", resolution_override=(dt.datetime(2026, 10, 3, 10, 0),))
    story(clean.key, "DC11", story_points=None)
    story(clean.key, "DC13", team=None)
    story(clean.key, "DC14", team="Another team")
    story(clean.key, "DC15", sprints=())
    story(clean.key, "DC16", sprints=CLOSED + (ACTIVE,))
    story(clean.key, "DC27", key_changes=(dt.datetime(2026, 10, 5, 10, 0),))
    story(clean.key, "DC29", pi="PI#1-2026")
    story(None, "DC01")
    story(story(clean.key).key, "DC02")                             # the parent is a story: wrong level
    wrong = issue(key(alpha), project=alpha, moves=CHAIN[:2], parent=clean.key, story_points=3, team=team_of(alpha),
                  sprints=(ACTIVE,))
    add(dataclasses.replace(wrong, changes=(wrong.changes[0], dataclasses.replace(wrong.changes[1], from_status="Planifié"))),
        ("DC06",))
    task = story(clean.key, type="Task")                            # sub-tasks under a story and under a task: DC03
    story(story(clean.key).key, "DC03", type="Sub-task", story_points=None)
    story(task.key, "DC03", type="Sub-task", story_points=None)
    story(story(clean.key, project=beta).key, "DC12", project=beta, type="Sub-task")        # points on a sub-task
    late = story(clean.key, "DC22", sprints=(dataclasses.replace(ACTIVE, start=dt.datetime(2026, 10, 30)),))
    story(clean.key, blocks=(late.key,))                            # its blocker ends (Nov 9) after it starts (Oct 30)
    unplanned = story(clean.key, "DC23", project=beta, moves=(), sprints=())                 # other team, no sprint
    story(clean.key, blocks=(unplanned.key,))
    for check_id, fields in (("DC17", dict(pi=None)), ("DC19", dict(target_end=None)),
                             ("DC20", dict(target_end=dt.date(2026, 8, 1))), ("DC21", dict(business_value=5)),
                             ("DC26", dict(fix_versions=()))):
        story(feature(check_id, **fields).key)                      # its own clean child: no DC24
    feature("DC24")
    story(feature("DC18").key, sprints=(BEFORE_PI,))                # a child whose sprint starts before the PI
    story(feature().key, "DC20", **{**FEATURE_DATES, "target_end": dt.date(2027, 2, 1)})    # dates after its parent's
    return dict(cases)


def with_anomalies(dataset: Dataset, projects: tuple[str, ...]) -> tuple[Dataset, dict[str, frozenset[str]]]:
    """The dataset plus the cases, and {key: expected check ids} for the cases."""
    cases = known_anomalies(projects)
    extended = dataclasses.replace(dataset, issues=dataset.issues + tuple(cases))
    return extended, {case.key: ids for case, ids in cases.items()}
