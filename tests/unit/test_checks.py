"""Data checks DC06..DC10 and DC28: the Python reference, the known anomalies of the fake Jira and the plan for Excel."""
import datetime as dt

import pytest

from jira2dashboard.flow import blueprint as B
from jira2dashboard.flow import formulas as F
from jira2dashboard.flow.checks import CHECKS, IDS, Moves, issue_checks
from jira2dashboard.flow.model import Change, FlowIssue, Rules
from jira2dashboard.mockjira.anomalies import known_anomalies, with_anomalies
from jira2dashboard.mockjira.flow_adapter import to_flow_issues
from jira2dashboard.mockjira.generator import generate
from jira2dashboard.params.model import project_keys

AS_OF = dt.datetime(2026, 11, 3, 18, 0)


@pytest.fixture(scope="module")
def rules():
    return Rules.from_params()


@pytest.fixture(scope="module")
def moves():
    return Moves.from_params()


def flow_issue(*moves_: tuple[str, str], status: str | None = None, created="2026-10-01T09:00", initial="Nouveau"):
    """Hand-made FlowIssue of the Team workflow: moves = (("2026-10-02T10:00", "En analyse"), ...)."""
    changes, previous = [], initial
    for n, (at, to) in enumerate(moves_, start=1):
        changes.append(Change(str(n), dt.datetime.fromisoformat(at), previous, to))
        previous = to
    return FlowIssue("T-1", "Alpha", "Story", "Autre", "Team", dt.datetime.fromisoformat(created),
                     status or previous, tuple(changes))


def test_the_checks_have_a_column_in_itemflow_and_a_row_in_the_plan():
    assert IDS == ("DC06", "DC07", "DC08", "DC09", "DC10", "DC28")
    itemflow = [name for name, _formula in F.ITEMFLOW_COLUMNS]
    assert all(i in itemflow for i in IDS)
    assert [r[0] for r in B.check_rows()] == list(IDS)
    assert all(r[1] and r[2] for r in B.check_rows())
    assert len({c.id for c in CHECKS}) == len(CHECKS)


def test_a_clean_history_fails_nothing(rules, moves):
    assert issue_checks(flow_issue(("2026-10-02T10:00", "En analyse"), ("2026-10-03T10:00", "Prêt")), AS_OF, rules, moves) == frozenset()


def test_an_issue_without_change_fails_nothing(rules, moves):
    assert issue_checks(flow_issue(), AS_OF, rules, moves) == frozenset()


def test_a_change_from_the_wrong_status_is_dc06(rules, moves):
    issue = flow_issue(("2026-10-02T10:00", "En analyse"), ("2026-10-03T10:00", "Prêt"))
    wrong = issue.changes[:1] + (Change("2", issue.changes[1].at, "Planifié", "Prêt"),)
    assert issue_checks(FlowIssue(**{**issue.__dict__, "changes": wrong}), AS_OF, rules, moves) == {"DC06"}


def test_a_current_status_that_the_history_does_not_reach_is_dc06(rules, moves):
    issue = flow_issue(("2026-10-02T10:00", "En analyse"), status="Prêt")
    assert issue_checks(issue, AS_OF, rules, moves) == {"DC06"}              # the history ends in En analyse, the issue says Prêt


def test_a_change_before_the_creation_is_dc07(rules, moves):
    issue = flow_issue(("2026-09-30T10:00", "En analyse"), ("2026-10-02T10:00", "Prêt"))
    assert issue_checks(issue, AS_OF, rules, moves) == {"DC07"}


def test_a_status_missing_from_workflows_is_dc08_and_a_move_that_is_not_listed_is_dc28(rules, moves):
    assert issue_checks(flow_issue(("2026-10-02T10:00", "En analyse"), ("2026-10-03T10:00", "Statut inconnu")),
                        AS_OF, rules, moves) == {"DC08", "DC28"}


def test_done_without_a_start_status_is_dc09(rules, moves):
    assert issue_checks(flow_issue(("2026-10-02T10:00", "Terminé")), AS_OF, rules, moves) == {"DC09", "DC28"}


def test_a_return_from_done_is_dc10(rules, moves):
    chain = [("2026-10-02T10:00", "En analyse"), ("2026-10-03T10:00", "Prêt"), ("2026-10-04T10:00", "Planifié"),
             ("2026-10-05T10:00", "En développement"), ("2026-10-06T10:00", "En test"), ("2026-10-07T10:00", "À déployer"),
             ("2026-10-08T10:00", "Terminé")]
    assert issue_checks(flow_issue(*chain), AS_OF, rules, moves) == frozenset()
    assert issue_checks(flow_issue(*chain, ("2026-10-09T10:00", "En développement")), AS_OF, rules, moves) == {"DC10", "DC28"}


def test_the_free_status_is_reachable_from_anywhere_but_not_left_freely(rules, moves):
    free = flow_issue(("2026-10-02T10:00", "En analyse"), ("2026-10-03T10:00", "Dépriorisé"))
    assert issue_checks(free, AS_OF, rules, moves) == frozenset()
    assert issue_checks(flow_issue(("2026-10-02T10:00", "Dépriorisé"), ("2026-10-03T10:00", "Prêt")), AS_OF, rules, moves) == {"DC28"}


def test_moves_follow_the_sheet_transitions_per_workflow(moves):
    assert moves.allowed("Team", "Nouveau", "En analyse")
    assert not moves.allowed("Team", "Nouveau", "Prêt")
    assert moves.allowed("Feature", "En développement", "En test / recette")
    assert not moves.allowed("Team", "En développement", "En test / recette")


def test_the_known_anomalies_fail_exactly_the_checks_they_were_written_for(rules, moves):
    cases = known_anomalies(project_keys()[1])
    assert len(cases) == 6 and any(not ids for ids in cases.values())          # five anomalies and one clean control
    dataset, expected = with_anomalies(generate(seed=3, issues=40, projects=project_keys()), project_keys()[1])
    by_key = {i.key: issue_checks(i, AS_OF, rules, moves) for i in to_flow_issues(dataset)}
    assert {key: by_key[key] for key in expected} == expected
    assert set().union(*expected.values()) == set(IDS)                         # every check is exercised by a case


def test_the_generated_issues_fail_no_check(rules, moves):
    dataset = generate(seed=11, issues=400, projects=project_keys())
    failed = {i.key: issue_checks(i, dataset.as_of, rules, moves) for i in to_flow_issues(dataset)}
    assert {k: v for k, v in failed.items() if v} == {}
