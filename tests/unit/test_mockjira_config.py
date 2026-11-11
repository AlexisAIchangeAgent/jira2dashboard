"""The Jira configuration answers of the fake Jira (mockjira/config.py): built from the dataset and the parameters."""
import pytest

from jira2dashboard.mockjira import config
from jira2dashboard.mockjira.errors import ApiError
from jira2dashboard.mockjira.generator import generate
from jira2dashboard.params.model import load_all

SHEETS = {s.name: s for s in load_all()}
DS = generate(seed=11, issues=250)                    # projects PRG (the Program), TA, TB


def statuses_of(workflow):
    return [row[2] for row in SHEETS["Workflows"].rows if row[0] == workflow]


def project(key):
    return config.project_search(DS, {"keys": [key], "expand": ["issueTypes"]})["values"][0]


def preview(key):
    found = project(key)
    return found, config.workflows_preview(DS, {"projectId": found["id"],
                                                "issueTypeIds": [t["id"] for t in found["issueTypes"]]})


def names(answer):
    """statusReference -> status name, inside one answer (a reference means nothing outside it)."""
    return {s["statusReference"]: s["name"] for s in answer["statuses"]}


def test_project_search_filters_by_keys_and_gives_ids_and_issue_types():
    page = config.project_search(DS, {"keys": ["TA", "PRG"], "expand": ["issueTypes"]})
    assert [p["key"] for p in page["values"]] == ["PRG", "TA"]
    assert (page["isLast"], page["total"], page["startAt"]) == (True, 2, 0)
    ids = [p["id"] for p in page["values"]]
    assert len(set(ids)) == 2 and all(i.isdigit() for i in ids)
    types = {p["key"]: [t["name"] for t in p["issueTypes"]] for p in page["values"]}
    assert types["PRG"] == ["Feature"]
    assert set(types["TA"]) == {i.type for i in DS.issues if i.project == "TA"}


def test_project_search_pages_by_offset_and_gives_issue_types_only_on_request():
    first = config.project_search(DS, {"maxResults": ["2"]})
    assert [p["key"] for p in first["values"]] == ["PRG", "TA"] and first["isLast"] is False
    assert "issueTypes" not in first["values"][0]
    second = config.project_search(DS, {"maxResults": ["2"], "startAt": ["2"]})
    assert [p["key"] for p in second["values"]] == ["TB"] and second["isLast"] is True


def test_issue_types_carry_the_subtask_flag_and_the_hierarchy_level():
    types = {t["name"]: t for t in project("TA")["issueTypes"]}
    assert (types["Sub-task"]["subtask"], types["Sub-task"]["hierarchyLevel"]) == (True, -1)
    assert (types["Story"]["subtask"], types["Story"]["hierarchyLevel"]) == (False, 0)
    assert project("PRG")["issueTypes"][0]["hierarchyLevel"] == 1


def test_project_statuses_per_issue_type_follow_the_workflow_of_the_type():
    answer = config.project_statuses(DS, "TA")
    assert {t["name"] for t in answer} == {t["name"] for t in project("TA")["issueTypes"]}
    for issue_type in answer:
        assert [s["name"] for s in issue_type["statuses"]] == statuses_of("Team")
    categories = {s["name"]: s["statusCategory"]["key"] for s in answer[0]["statuses"]}
    assert (categories["Nouveau"], categories["Prêt"], categories["Terminé"]) == ("new", "indeterminate", "done")
    assert [s["name"] for s in config.project_statuses(DS, "PRG")[0]["statuses"]] == statuses_of("Feature")


def test_project_statuses_of_an_unknown_project_is_404():
    with pytest.raises(ApiError) as err:
        config.project_statuses(DS, "NOPE")
    assert err.value.status == 404


def test_preview_gives_each_workflow_with_the_issue_types_of_the_project_that_use_it():
    found, answer = preview("TA")
    assert [w["name"] for w in answer["workflows"]] == ["Team"]
    assert answer["workflows"][0]["queryContext"] == [{"project": found["id"],
                                                       "issueTypes": [t["id"] for t in found["issueTypes"]]}]
    assert [w["name"] for w in preview("PRG")[1]["workflows"]] == ["Feature"]


def test_preview_statuses_and_directed_transitions_are_the_parameters():
    _found, answer = preview("TA")
    ref, workflow = names(answer), answer["workflows"][0]
    assert [ref[s["statusReference"]] for s in workflow["statuses"]] == statuses_of("Team")
    moves = [(ref[link["fromStatusReference"]], ref[t["toStatusReference"]])
             for t in workflow["transitions"] if t["type"] == "DIRECTED" for link in t["links"]]
    assert sorted(moves) == sorted((r[1], r[2]) for r in SHEETS["Transitions"].rows if r[0] == "Team")
    several = [t for t in workflow["transitions"] if len(t["links"]) > 1]
    assert several, "a status reached from several statuses is one transition with several links, as in Jira"


def test_preview_global_transitions_lead_to_the_free_statuses_and_initial_to_the_first():
    _found, answer = preview("TA")
    ref, transitions = names(answer), answer["workflows"][0]["transitions"]
    team = [r for r in SHEETS["Workflows"].rows if r[0] == "Team"]
    assert [ref[t["toStatusReference"]] for t in transitions if t["type"] == "GLOBAL"] == [r[2] for r in team if r[6] == 1]
    assert [ref[t["toStatusReference"]] for t in transitions if t["type"] == "INITIAL"] == [r[2] for r in team if r[1] == 1]
    assert all(t["links"] == [] for t in transitions if t["type"] != "DIRECTED")


def test_preview_categories_use_the_preview_enum_and_references_are_not_ids():
    _found, answer = preview("TA")
    categories = {s["name"]: s["statusCategory"] for s in answer["statuses"]}
    assert (categories["Nouveau"], categories["En développement"], categories["Annulé"]) == ("TODO", "IN_PROGRESS", "DONE")
    assert all(s["statusReference"] != s["id"] for s in answer["statuses"])


@pytest.mark.parametrize("body, status", [
    ({"issueTypeIds": ["10002"]}, 400),                                          # no projectId
    ({"projectId": "PID"}, 400),                                                 # no lookup term
    ({"projectId": "99999", "issueTypeIds": ["10002"]}, 404),                    # unknown project
    ({"projectId": "PID", "issueTypeIds": [str(n) for n in range(26)]}, 400),    # more than 25 issue types
])
def test_preview_checks_its_body(body, status):
    body = {k: (project("TA")["id"] if v == "PID" else v) for k, v in body.items()}
    with pytest.raises(ApiError) as err:
        config.workflows_preview(DS, body)
    assert err.value.status == status


def test_create_meta_gives_the_fields_of_the_create_screen_with_paging():
    found = project("TA")
    type_id = found["issueTypes"][0]["id"]
    first = config.create_meta_fields(DS, "TA", type_id, {"startAt": ["0"], "maxResults": ["5"]})
    assert (first["startAt"], first["maxResults"], len(first["fields"])) == (0, 5, 5)
    assert first["total"] == len(DS.fields.catalog())
    last = config.create_meta_fields(DS, "TA", type_id, {"startAt": [str(first["total"] - 1)]})
    assert len(last["fields"]) == 1
    required = {f["fieldId"] for f in first["fields"] + last["fields"] if f["required"]}
    assert required <= set(config.CREATE_META_REQUIRED)


def test_create_meta_of_an_unknown_project_or_an_absent_issue_type_is_404():
    with pytest.raises(ApiError) as err:
        config.create_meta_fields(DS, "NOPE", "10002", {})
    assert err.value.status == 404
    with pytest.raises(ApiError) as err:
        config.create_meta_fields(DS, "TA", "99999", {})
    assert err.value.status == 404
