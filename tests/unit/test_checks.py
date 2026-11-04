"""Data checks: the Python reference (flow/checks.py), the known anomalies of the fake Jira and the plan for Excel."""
import dataclasses
import datetime as dt

import pytest

from jira2dashboard.flow import blueprint as B
from jira2dashboard.flow import check_formulas as C
from jira2dashboard.flow import formulas as F
from jira2dashboard.flow.checks import (CHECKS, FILL_FIELDS, IDS, PER_ISSUE, CheckSettings, Moves, fill_rates, run_checks,
                                         valid_pi)
from jira2dashboard.flow.model import Change, FlowIssue, Rules
from jira2dashboard.mockjira.anomalies import known_anomalies, with_anomalies
from jira2dashboard.mockjira.flow_adapter import to_flow_issues
from jira2dashboard.mockjira.generator import generate
from jira2dashboard.params.model import project_keys

AS_OF = dt.datetime(2026, 11, 3, 18, 0)
D = dt.datetime
FIELD_PARAMETERS = ("StoryPointsField", "BusinessValueField", "PIField", "TeamField", "SprintField", "TargetStartField",
                    "TargetEndField")
CONFIGURED = CheckSettings(
    parent_types=frozenset({"Story", "Task"}), feature_types=frozenset({"Feature"}), subtask_types=frozenset({"Sub-task"}),
    points_types=frozenset({"Story"}), fields={name: f"customfield_{n}" for n, name in enumerate(FIELD_PARAMETERS)},
    fill_threshold=90, pis={"PI#01-2026": (D(2026, 6, 1), D(2026, 12, 31))})
UNCONFIGURED = dataclasses.replace(CONFIGURED, fields={}, pis={})


def check(check_id):
    return next(c for c in CHECKS if c.id == check_id)


@pytest.fixture(scope="module")
def rules():
    return Rules.from_params()


@pytest.fixture(scope="module")
def moves():
    return Moves.from_params()


def story(key="S-1", **changes):
    """A clean Team issue: a parent, a team, a sprint, story points. `changes` override fields."""
    base = dict(key=key, team="Alpha", type="Story", work_type="Autre", workflow="Team", created=D(2026, 10, 1, 9),
                status="En analyse", status_category="indeterminate", parent="F-1", story_points=3, team_value="Alpha",
                sprint_state="active", sprint_start=D(2026, 10, 26), sprint_end=D(2026, 11, 9), sprint_count=1,
                changes=(Change("1", D(2026, 10, 2, 10), "Nouveau", "En analyse"),))
    return FlowIssue(**{**base, **changes})


def feature(key="F-1", **changes):
    base = dict(key=key, team="ART", type="Feature", work_type="Feature", workflow="Feature", created=D(2026, 9, 1, 9),
                status="En analyse", status_category="indeterminate", team_value="ART", pi="PI#01-2026",
                fix_versions=("R1",), target_start=D(2026, 9, 1), target_end=D(2026, 12, 31),
                changes=(Change("1", D(2026, 9, 2, 10), "Nouveau", "En analyse"),))
    return FlowIssue(**{**base, **changes})


def failed(rules, moves, issues, settings=CONFIGURED):
    return run_checks(issues, AS_OF, rules, moves, settings)


def test_the_catalogue_covers_dc01_to_dc29_once_and_each_per_issue_check_has_a_formula():
    assert IDS == tuple(f"DC{n:02d}" for n in range(1, 30))
    assert "DC25" not in PER_ISSUE and len(PER_ISSUE) == 28
    assert all(i in dict(F.ITEMFLOW_COLUMNS) for i in PER_ISSUE)
    assert [r[0] for r in B.check_rows()] == list(IDS)
    assert all(r[1] and r[2] for r in B.check_rows()) and all(len(r) == len(B.CHECKS_HEADER) for r in B.check_rows())
    assert {r[0]: r[6] for r in B.check_rows()}["DC25"] == "FillRates"
    assert tuple(C.FLAGS) == PER_ISSUE


def test_a_clean_feature_with_a_clean_story_fails_nothing(rules, moves):
    result = failed(rules, moves, [feature(), story()])
    assert result == {"F-1": frozenset(), "S-1": frozenset()}


CASES = [
    ("DC01", {"parent": None}, ()),
    ("DC02", {"parent": "S-9"}, (story("S-9"),)),
    ("DC04", {"status_category": "done"}, ()),
    ("DC05", {"resolved": D(2026, 10, 3)}, ()),
    ("DC11", {"story_points": None}, ()),
    ("DC13", {"team_value": None}, ()),
    ("DC14", {"team_value": "Beta"}, ()),
    ("DC15", {"sprint_state": "", "sprint_count": 0, "sprint_start": None, "sprint_end": None}, ()),
    ("DC16", {"sprint_count": 3}, ()),
    ("DC27", {"key_changes": 1}, ()),
    ("DC29", {"pi": "PI#1-2026"}, ()),
]


@pytest.mark.parametrize("check_id,change,extra", CASES, ids=[c[0] for c in CASES])
def test_a_clean_story_with_one_field_changed_fails_exactly_that_check(rules, moves, check_id, change, extra):
    result = failed(rules, moves, [feature(), *extra, story(**change)])
    assert result["S-1"] == {check_id}


def test_dc02_needs_a_parent_that_is_in_the_extract_and_is_not_a_feature(rules, moves):
    assert failed(rules, moves, [story(parent="Z-1")])["S-1"] == frozenset()          # parent not extracted: not judged
    assert "DC02" not in failed(rules, moves, [feature(), story()])["S-1"]


def test_dc03_flags_the_subtasks_of_a_team_whose_parents_are_of_two_types(rules, moves):
    parents = [story("S-1"), story("T-1", type="Task")]
    subtasks = [story("U-1", type="Sub-task", parent="S-1", story_points=None, sprint_state="", sprint_count=0),
                story("U-2", type="Sub-task", parent="T-1", story_points=None, sprint_state="", sprint_count=0)]
    result = failed(rules, moves, [feature(), *parents, *subtasks])
    assert "DC03" in result["U-1"] and "DC03" in result["U-2"] and "DC03" not in result["S-1"]
    assert "DC03" not in failed(rules, moves, [feature(), *parents, subtasks[0]])["U-1"]


def test_dc12_dc21_dc24_dc26_are_about_the_type(rules, moves):
    sub = story("U-1", type="Sub-task", parent="S-1", story_points=2)
    lone = feature(business_value=5, fix_versions=())
    result = failed(rules, moves, [lone, story(parent=None), sub])
    assert "DC12" in result["U-1"]
    assert {"DC21", "DC24", "DC26"} <= result["F-1"]
    assert "DC24" not in failed(rules, moves, [lone, story()])["F-1"]


def test_dc17_and_dc29_use_the_pi_format(rules, moves):
    assert valid_pi("PI#01-2026") and not valid_pi("PI#1-2026") and not valid_pi("pi#01-2026") and not valid_pi(None)
    assert not valid_pi("Champ_customefield_10300") and not valid_pi("PI#01-202a")
    placeholder = feature(pi="Champ_customefield_10300")
    assert {"DC17", "DC29"} <= failed(rules, moves, [placeholder])["F-1"]
    assert "DC17" in failed(rules, moves, [feature(pi=None)])["F-1"]
    assert "DC29" not in failed(rules, moves, [feature(pi=None)])["F-1"]


def test_dc18_compares_the_sprints_of_the_children_with_the_dates_of_the_pi(rules, moves):
    inside = story(sprint_start=D(2026, 7, 1), sprint_end=D(2026, 7, 15))
    outside = story("S-2", sprint_start=D(2026, 5, 20), sprint_end=D(2026, 6, 3))
    assert "DC18" not in failed(rules, moves, [feature(), inside])["F-1"]
    assert "DC18" in failed(rules, moves, [feature(), inside, outside])["F-1"]
    assert "DC18" not in failed(rules, moves, [feature(pi="PI#09-2026"), outside])["F-1"]      # a PI that PIs does not know


def test_dc19_and_dc20_use_the_target_dates(rules, moves):
    assert "DC19" in failed(rules, moves, [feature(target_end=None)])["F-1"]
    backlog = feature(target_end=None, status="Nouveau", status_category="new", changes=())
    assert "DC19" not in failed(rules, moves, [backlog])["F-1"]
    assert "DC20" in failed(rules, moves, [feature(target_end=D(2026, 8, 1))])["F-1"]
    after = story(target_start=D(2026, 10, 1), target_end=D(2027, 2, 1))
    assert "DC20" in failed(rules, moves, [feature(), after])["S-1"]
    within = story(target_start=D(2026, 10, 1), target_end=D(2026, 11, 1))
    assert "DC20" not in failed(rules, moves, [feature(), within])["S-1"]


def test_dc22_and_dc23_come_from_the_blocking_links(rules, moves):
    blocker = story("S-1", blocks=("S-2",), sprint_start=D(2026, 10, 26), sprint_end=D(2026, 11, 9))
    late = story("S-2", sprint_start=D(2026, 10, 30), sprint_end=D(2026, 11, 13))
    later = story("S-3", sprint_start=D(2026, 11, 9), sprint_end=D(2026, 11, 23))
    assert "DC22" in failed(rules, moves, [feature(), blocker, late])["S-2"]
    assert "DC22" not in failed(rules, moves, [feature(), dataclasses.replace(blocker, blocks=("S-3",)), later])["S-3"]
    beta = story("S-4", team="Beta", team_value="Beta", sprint_state="", sprint_count=0, sprint_start=None, sprint_end=None)
    cross = failed(rules, moves, [feature(), dataclasses.replace(blocker, blocks=("S-4",)), beta])
    assert "DC23" in cross["S-4"] and "DC23" not in cross["S-1"]
    unplanned_same_team = dataclasses.replace(late, sprint_count=0, sprint_start=None, sprint_end=None)
    assert "DC23" not in failed(rules, moves, [feature(), blocker, unplanned_same_team])["S-2"]


def test_a_check_that_needs_a_field_flags_nobody_while_the_field_is_not_configured(rules, moves):
    broken = story(story_points=None, team_value="Beta", sprint_count=3, pi="x")
    assert {"DC11", "DC14", "DC16", "DC29"} <= failed(rules, moves, [feature(), broken])["S-1"]
    assert failed(rules, moves, [feature(), broken], UNCONFIGURED)["S-1"] == frozenset()
    assert {c.id for c in CHECKS if not UNCONFIGURED.ready(c)} == {c.id for c in CHECKS if c.needs}
    assert all(CONFIGURED.ready(c) for c in CHECKS)
    assert UNCONFIGURED.ready(check("DC01")) and not UNCONFIGURED.ready(check("DC18"))


def test_pis_without_rows_make_dc18_not_computable_even_with_the_fields():
    no_pis = dataclasses.replace(CONFIGURED, pis={})
    assert not no_pis.ready(check("DC18")) and no_pis.ready(check("DC17"))


def test_the_status_history_checks_still_work(rules, moves):
    def with_history(key, *moves_, **fields):
        changes, previous = [], "Nouveau"
        for n, (at, to) in enumerate(moves_, start=1):
            changes.append(Change(str(n), at, previous, to))
            previous = to
        return story(key, changes=tuple(changes), status=previous, **fields)

    skipped = with_history("S-1", (D(2026, 10, 2, 10), "Terminé"), status_category="done", resolved=D(2026, 10, 2, 10))
    assert {"DC09", "DC28"} <= failed(rules, moves, [feature(), skipped])["S-1"]
    unknown = with_history("S-2", (D(2026, 10, 2, 10), "En analyse"), (D(2026, 10, 3, 10), "Statut inconnu"))
    assert {"DC08", "DC28"} <= failed(rules, moves, [feature(), unknown])["S-2"]
    before = with_history("S-3", (D(2026, 9, 30, 10), "En analyse"))
    assert "DC07" in failed(rules, moves, [feature(), before])["S-3"]
    wrong_from = story("S-4", status="Prêt", changes=(Change("1", D(2026, 10, 2, 10), "Nouveau", "En analyse"),
                                                      Change("2", D(2026, 10, 3, 10), "Planifié", "Prêt")))
    assert "DC06" in failed(rules, moves, [feature(), wrong_from])["S-4"]
    free = with_history("S-5", (D(2026, 10, 2, 10), "Dépriorisé"))
    assert "DC28" not in failed(rules, moves, [feature(), free])["S-5"]
    chain = [(D(2026, 10, 1 + n, 10), to) for n, to in enumerate(
        ["En analyse", "Prêt", "Planifié", "En développement", "En test", "À déployer", "Terminé"], start=1)]
    reopened = with_history("S-6", *chain, (D(2026, 10, 12, 10), "En développement"))
    assert {"DC10", "DC28"} <= failed(rules, moves, [feature(), reopened])["S-6"]
    assert "DC10" not in failed(rules, moves, [feature(), with_history("S-7", *chain, status_category="done",
                                                                       resolved=D(2026, 10, 8, 10))])["S-7"]


def test_fill_rates_count_filled_values_per_field_team_and_type():
    issues = [story("S-1"), story("S-2", story_points=None), story("S-3", story_points=None)]
    rates = fill_rates(issues, CONFIGURED, ["Alpha", "Beta"], ["Story"])
    by = {(r.field, r.team): r for r in rates}
    assert by[("StoryPoints", "Alpha")].count == 3 and by[("StoryPoints", "Alpha")].filled == 1
    assert by[("StoryPoints", "Alpha")].below and not by[("Parent", "Alpha")].below
    assert by[("StoryPoints", "Beta")].count == 0 and not by[("StoryPoints", "Beta")].below
    assert len(rates) == len(FILL_FIELDS) * 2
    assert {r.field for r in fill_rates(issues, UNCONFIGURED, ["Alpha"], ["Story"])} == {n for n, need in FILL_FIELDS if not need}


def test_moves_follow_the_sheet_transitions_per_workflow(moves):
    assert moves.allowed("Team", "Nouveau", "En analyse")
    assert not moves.allowed("Team", "Nouveau", "Prêt")
    assert moves.allowed("Feature", "En développement", "En test / recette")
    assert not moves.allowed("Team", "En développement", "En test / recette")


def test_the_settings_come_from_the_config_and_pis_sheets():
    settings = CheckSettings.from_params()
    assert settings.parent_types == {"Story", "Task"} and settings.feature_types == {"Feature"}
    assert settings.subtask_types == {"Sub-task"} and settings.points_types == {"Story"}
    assert settings.fill_threshold == 90 and settings.pis == {}
    assert set(settings.fields) == set(FIELD_PARAMETERS) and not any(settings.fields.values())   # client values stay empty


def test_the_known_anomalies_fail_exactly_the_checks_they_were_written_for(rules, moves):
    assert any(not ids for ids in known_anomalies(project_keys()).values())          # at least one clean control
    dataset, expected = with_anomalies(generate(seed=3, issues=40, projects=project_keys()), project_keys())
    result = run_checks(to_flow_issues(dataset), AS_OF, rules, moves, CONFIGURED)
    assert {key: result[key] for key in expected} == expected
    assert set().union(*expected.values()) == set(PER_ISSUE)                          # every per-issue check has a case


def test_the_formulas_name_only_columns_that_exist():
    assert {name for name, _need in C.FILL_FIELD_NEEDS} <= set(F.TABLES["ItemFlow"])
    assert set(C.RAW_FIELDS) <= set(F.ISSUES) and len(F.ISSUES) == len(set(F.ISSUES))
