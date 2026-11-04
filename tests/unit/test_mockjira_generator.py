"""Generated volume data: deterministic and consistent with the workflow parameters."""
import datetime as dt
from collections import Counter

import pytest

from jira2dashboard.mockjira.generator import generate
from jira2dashboard.params.model import load_all


@pytest.fixture(scope="module")
def dataset():
    return generate(seed=7, issues=400)


@pytest.fixture(scope="module")
def sheets():
    return {s.name: s for s in load_all()}


def test_same_seed_same_data_other_seed_other_data():
    assert generate(seed=3, issues=50) == generate(seed=3, issues=50)
    assert generate(seed=3, issues=50) != generate(seed=4, issues=50)


def test_requested_number_of_issues_with_unique_keys_and_ids(dataset):
    assert len(dataset.issues) == 400
    assert len({i.key for i in dataset.issues}) == 400 and len({i.id for i in dataset.issues}) == 400


def test_three_projects_with_the_expected_types(dataset):
    assert dataset.projects() == ["PRG", "TA", "TB"]
    types = {p: {i.type for i in dataset.issues if i.project == p} for p in dataset.projects()}
    assert types["PRG"] == {"Feature"}
    assert "Feature" not in types["TA"] and {"Story", "Task", "Bug"} <= types["TA"]


def test_every_move_follows_a_transition_of_the_parameters(dataset, sheets):
    allowed = {(wf, a, b) for wf, a, b, _kind in sheets["Transitions"].rows}
    free = {wf: next(r[2] for r in sheets["Workflows"].rows if r[0] == wf and r[6] == 1) for wf in ("Team", "Feature")}
    for issue in dataset.issues:
        previous = issue.initial_status
        for change in issue.changes:
            assert change.from_status == previous
            ok = (issue.workflow, previous, change.to_status) in allowed or change.to_status == free[issue.workflow]
            assert ok, (issue.key, previous, change.to_status)
            previous = change.to_status


def test_histories_are_in_time_order_inside_the_observation_window(dataset):
    for issue in dataset.issues:
        times = [issue.created] + [c.at for c in issue.changes]
        assert times == sorted(times), issue.key
        assert times[-1] <= dataset.as_of


def test_features_use_the_feature_workflow_and_others_the_team_workflow(dataset):
    assert {i.workflow for i in dataset.issues if i.type == "Feature"} == {"Feature"}
    assert {i.workflow for i in dataset.issues if i.type != "Feature"} == {"Team"}


def test_the_cases_flow_metrics_must_handle_are_present(dataset):
    statuses = Counter(i.status for i in dataset.issues)
    assert statuses["Terminé"] > 0 and statuses["Annulé"] > 0 and statuses["Nouveau"] > 0
    restarted = [i for i in dataset.issues
                 if any(a.to_status == "Dépriorisé" and b.to_status == "Nouveau"
                        for a, b in zip(i.changes, i.changes[1:]))]
    stepped_back = [i for i in dataset.issues
                    if any(a.from_status == "Prêt" and a.to_status == "En analyse" for a in i.changes)]
    assert restarted and stepped_back


def test_parents_exist_and_have_the_right_type(dataset):
    by_key = {i.key: i for i in dataset.issues}
    for issue in dataset.issues:
        if issue.parent:
            assert issue.parent in by_key
            assert by_key[issue.parent].type == ("Feature" if issue.type == "Story" else "Story")


def test_project_keys_are_a_parameter_the_first_one_is_the_program():
    ds = generate(seed=2, issues=120, projects=("TTT", "AAA", "EEE", "FFFF"))
    assert ds.projects() == ["AAA", "EEE", "FFFF", "TTT"]
    assert {i.type for i in ds.issues if i.project == "TTT"} == {"Feature"}
    assert all(i.type != "Feature" for i in ds.issues if i.project != "TTT")


def test_a_program_and_at_least_one_team_are_needed():
    with pytest.raises(ValueError, match="Program"):
        generate(issues=5, projects=("ONLY",))


def test_custom_field_ids_are_a_parameter():
    from jira2dashboard.mockjira.dataset import FieldIds
    ids = FieldIds(story_points="customfield_20001", business_value="customfield_20002", pi="customfield_20003")
    ds = generate(seed=2, issues=30, fields=ids)
    assert ds.fields == ids
    assert {f["id"] for f in ds.fields.catalog()} >= {"customfield_20001", "customfield_20002", "customfield_20003"}


def test_issues_are_created_between_since_and_as_of(dataset):
    assert all(dt.datetime(2026, 6, 1) <= i.created <= dataset.as_of for i in dataset.issues)
