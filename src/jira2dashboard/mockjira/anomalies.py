"""Issues with known data anomalies, to test the data checks (flow/checks.py) on the fake Jira.

Each case is written by hand with the ids of the checks it must trigger. The generated issues follow the allowed
moves and have a parent when their type needs one, so they trigger none of the status-history checks.
"""
import dataclasses
import datetime as dt

from jira2dashboard.mockjira.dataset import Dataset, Issue, issue

FIRST_NUMBER = 9001                       # keys of the cases: <project>-9001 ... (the generated ones stay far below)


def known_anomalies(project: str) -> dict[Issue, frozenset[str]]:
    """The cases, in a team `project`, each with the ids of the checks that it fails."""
    def key(n: int) -> str:
        return f"{project}-{FIRST_NUMBER + n}"

    chain = (("2026-10-02T10:00", "En analyse"), ("2026-10-03T10:00", "Prêt"), ("2026-10-04T10:00", "Planifié"),
             ("2026-10-05T10:00", "En développement"), ("2026-10-06T10:00", "En test"),
             ("2026-10-07T10:00", "À déployer"), ("2026-10-08T10:00", "Terminé"))
    parent = f"{project}-1"                                   # every case has a parent but the orphan

    def case(n: int, moves, **fields):
        return issue(key(n), moves=moves, parent=parent, **fields)

    skipped = case(0, (("2026-10-02T10:00", "Terminé"),))
    reopened = case(1, chain + (("2026-10-09T10:00", "En développement"),))
    unknown = case(2, chain[:1] + (("2026-10-03T10:00", "Statut inconnu"),))
    clean = case(3, chain[:2])
    wrong_from = case(4, chain[:2])
    wrong_from = dataclasses.replace(wrong_from, changes=(
        wrong_from.changes[0], dataclasses.replace(wrong_from.changes[1], from_status="Planifié")))
    before_created = case(5, (("2026-09-30T10:00", "En analyse"), ("2026-10-02T10:00", "Prêt")))
    orphan = issue(key(6), moves=chain[:2])
    unresolved = case(7, chain, resolution_override=(None,))
    resolved_early = case(8, chain[:2], resolution_override=(dt.datetime(2026, 10, 3, 10, 0),))
    return {skipped: frozenset({"DC09", "DC28"}), reopened: frozenset({"DC10", "DC28"}),
            unknown: frozenset({"DC08", "DC28"}), clean: frozenset(), wrong_from: frozenset({"DC06"}),
            before_created: frozenset({"DC07"}), orphan: frozenset({"DC01"}), unresolved: frozenset({"DC04"}),
            resolved_early: frozenset({"DC05"})}


def with_anomalies(dataset: Dataset, project: str) -> tuple[Dataset, dict[str, frozenset[str]]]:
    """The dataset plus the cases, and {key: expected check ids} for the cases."""
    cases = known_anomalies(project)
    extended = dataclasses.replace(dataset, issues=dataset.issues + tuple(cases))
    return extended, {case.key: ids for case, ids in cases.items()}
