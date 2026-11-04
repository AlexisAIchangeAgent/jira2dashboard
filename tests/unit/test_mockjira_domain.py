"""Fake Jira: dataset, JQL subset, and the pure API responses."""
import datetime as dt

import pytest

from jira2dashboard.mockjira import api, jql
from jira2dashboard.mockjira.dataset import Dataset, FieldIds, StatusCatalog, issue
from jira2dashboard.mockjira.errors import ApiError

TZ = 120
BASE = "http://mock"


# --- dataset ----------------------------------------------------------------------------------------

def test_issue_status_is_the_last_move_or_the_initial_status(small_dataset):
    assert small_dataset.by_key_or_id("TA-1").status == "Terminé"
    assert small_dataset.by_key_or_id("TA-2").status == "En analyse"
    assert small_dataset.by_key_or_id("TA-3").status == "Nouveau"


def test_issue_can_be_found_by_key_or_id(small_dataset):
    assert small_dataset.by_key_or_id("10004").key == "TB-1"
    assert small_dataset.by_key_or_id("nope") is None


def test_duplicate_keys_are_refused():
    with pytest.raises(ValueError, match="duplicate"):
        Dataset((issue("A-1"), issue("A-1", id="10009")), dt.datetime(2026, 11, 3))


def test_resolution_exists_only_for_done_and_removed_statuses(small_dataset):
    assert small_dataset.resolution(small_dataset.by_key_or_id("TA-1")) == ("Done", dt.datetime(2026, 10, 14, 10, 0))
    assert small_dataset.resolution(small_dataset.by_key_or_id("TB-1"))[0] == "Won't Do"
    assert small_dataset.resolution(small_dataset.by_key_or_id("TA-2")) is None


def test_status_catalog_comes_from_the_workflow_parameters():
    catalog = StatusCatalog.from_params()
    assert len(catalog.ids) == 11 and len(set(catalog.ids.values())) == 11     # 10 Team + "En test / recette"
    assert catalog.end_statuses == {"Team": "Terminé", "Feature": "Terminé"}
    assert catalog.jira_category("Team", "Terminé") == ("done", "Terminé")
    assert catalog.jira_category("Team", "En analyse") == ("indeterminate", "En cours")
    assert catalog.jira_category("Team", "Nouveau") == ("new", "À faire")


def test_field_catalog_lists_system_and_custom_fields():
    fields = {f["id"]: f for f in FieldIds().catalog()}
    assert fields["status"]["custom"] is False
    assert fields["customfield_10016"]["name"] == "Story point estimate" and fields["customfield_10016"]["custom"]


# --- JQL --------------------------------------------------------------------------------------------

def test_jql_project_in_with_order():
    query = jql.parse('project in ("TA","TB") ORDER BY key ASC')
    assert query == jql.Query(("TA", "TB"), "key", False)


def test_jql_single_project_and_descending_order():
    assert jql.parse("project = TA order by created desc") == jql.Query(("TA",), "created", True)


def test_unbounded_jql_is_refused_like_jira():
    for text in ("", "ORDER BY key ASC"):
        with pytest.raises(ApiError) as err:
            jql.parse(text)
        assert err.value.status == 400 and "Unbounded JQL queries are not allowed" in err.value.body["errorMessages"][0]


def test_unsupported_jql_is_refused_not_misread():
    with pytest.raises(ApiError, match="supports"):
        jql.parse("project = TA AND assignee = bob")
    with pytest.raises(ApiError, match="ORDER BY"):
        jql.parse("project = TA ORDER BY priority")
    with pytest.raises(ApiError, match="one project clause"):
        jql.parse("project = TA AND project = TB")
    with pytest.raises(ApiError, match="Unbounded"):
        jql.parse("status = Done")


def test_jql_with_a_filter_restriction_is_parsed_into_conditions():
    query = jql.parse('project in (TTT, "AAA") AND (updated >= -90d AND issuetype in (Story, "Sub-task")) '
                      'and status = "En analyse" ORDER BY created DESC')
    assert query.projects == ("TTT", "AAA") and query.order_by == "created" and query.descending
    assert query.conditions == (
        jql.Condition("updated", op=">=", bound="-90d"),
        jql.Condition("issuetype", values=("Story", "Sub-task")),
        jql.Condition("status", values=("En analyse",)))


def test_and_inside_quotes_or_parentheses_does_not_split():
    assert jql._split_and('a = "x AND y" AND b') == ['a = "x AND y"', "b"]
    assert jql._split_and("(a AND b) AND c") == ["(a AND b)", "c"]


def select_keys(ds, text):
    return [i["key"] for i in run_search(ds, jql=text, fields=["summary"])["issues"]]


def test_filter_on_issue_type_and_status(small_dataset):
    assert select_keys(small_dataset, "project in (TA,TB) AND issuetype = Story") == ["TA-1", "TB-1"]
    assert select_keys(small_dataset, 'project = TA AND (status in ("En analyse","Nouveau"))') == ["TA-2", "TA-3"]
    assert select_keys(small_dataset, "project = TA AND issuetype = story") == ["TA-1"]          # names ignore case


def test_filter_on_relative_dates_counts_from_the_as_of_date(small_dataset):
    # as of 2026-11-03T18:00: -25d is 2026-10-09T18:00; only TA-1 (updated 2026-10-14) is newer
    assert select_keys(small_dataset, "project in (TA,TB,PRG) AND updated >= -25d") == ["TA-1"]
    assert select_keys(small_dataset, "project in (TA,TB,PRG) AND updated >= -2w") == []
    assert len(select_keys(small_dataset, "project in (TA,TB,PRG) AND updated >= -200d")) == 5


def test_filter_on_absolute_dates(small_dataset):
    created_on_the_1st = 'project in (TA,TB,PRG) AND created >= "2026-10-01" AND created < 2026-10-02'
    assert select_keys(small_dataset, created_on_the_1st) == ["TA-1", "TA-2", "TA-3", "TB-1"]   # PRG-1 was created in August
    assert select_keys(small_dataset, 'project in (TA,TB,PRG) AND updated < "2026-10-09"') == ["PRG-1", "TA-2", "TA-3"]


def test_a_filter_can_wrap_the_whole_query_and_combine_with_order(small_dataset):
    assert select_keys(small_dataset, "(project = TA AND issuetype = Bug)") == ["TA-3"]
    assert select_keys(small_dataset, "project = TA AND updated >= -200d ORDER BY key DESC") == ["TA-3", "TA-2", "TA-1"]


# --- search -----------------------------------------------------------------------------------------

def run_search(ds, clamp=100, **params):
    return api.search(ds, params, tz=TZ, base_url=BASE, page_clamp=clamp)


def test_search_without_fields_returns_ids_only(small_dataset):
    body = run_search(small_dataset, jql='project in ("TA")')
    assert body["issues"] == [{"id": "10001"}, {"id": "10002"}, {"id": "10003"}]
    assert body["isLast"] is True and body["nextPageToken"] is None          # null on the last page, like Jira


def test_search_all_fields_with_names(small_dataset):
    body = run_search(small_dataset, jql="project = TA", fields=["*all"], expand="names")
    first = body["issues"][0]
    assert first["key"] == "TA-1" and first["self"] == f"{BASE}/rest/api/3/issue/10001"
    f = first["fields"]
    assert f["status"] == {"id": StatusCatalog.from_params().ids["Terminé"], "name": "Terminé",
                           "statusCategory": {"key": "done", "name": "Terminé"}}
    assert f["created"] == "2026-10-01T11:00:00.000+0200"           # 09:00 UTC shown at +0200
    assert f["resolutiondate"] == "2026-10-14T12:00:00.000+0200"
    assert f["resolution"] == {"name": "Done"} and f["parent"] == {"key": "PRG-1"}
    assert f["customfield_10016"] == 5 and f["customfield_10200"] == "Feature"
    assert body["names"]["status"] == "Status" and body["names"]["customfield_10016"] == "Story point estimate"


def test_empty_fields_are_present_as_null_so_columns_are_stable(small_dataset):
    f = run_search(small_dataset, jql="project = TA", fields=["*all"])["issues"][2]["fields"]
    assert f["resolution"] is None and f["assignee"] is None and f["parent"] is None and f["duedate"] is None


def test_search_with_a_field_list_returns_only_those_fields(small_dataset):
    body = run_search(small_dataset, jql="project = TA", fields=["summary", "status"])
    assert set(body["issues"][0]["fields"]) == {"summary", "status"}
    assert "names" not in body


def test_search_is_sorted_by_key_and_paged_by_token(small_dataset):
    seen, token = [], None
    for _ in range(10):
        body = run_search(small_dataset, jql='project in ("TB","TA","PRG")', fields=["summary"], maxResults=2,
                          nextPageToken=token)
        seen += [i["key"] for i in body["issues"]]
        token = body.get("nextPageToken")
        if body["isLast"]:
            break
    assert seen == ["PRG-1", "TA-1", "TA-2", "TA-3", "TB-1"]
    assert token is None


def test_page_size_is_clamped_by_the_server(small_dataset):
    body = run_search(small_dataset, jql='project in ("TA","TB")', fields=["summary"], maxResults=100, clamp=2)
    assert len(body["issues"]) == 2 and body["isLast"] is False and "nextPageToken" in body


def test_default_page_size_is_50(small_dataset):
    ds = Dataset(tuple(issue(f"A-{n}", id=str(20000 + n)) for n in range(1, 61)), dt.datetime(2026, 11, 3))
    body = run_search(ds, jql="project = A", fields=["summary"])
    assert len(body["issues"]) == 50 and "nextPageToken" in body


def test_token_of_another_query_is_refused(small_dataset):
    token = run_search(small_dataset, jql="project in (TA, TB)", maxResults=1)["nextPageToken"]
    with pytest.raises(ApiError, match="nextPageToken"):
        run_search(small_dataset, jql="project = TA", nextPageToken=token)
    with pytest.raises(ApiError, match="nextPageToken"):
        run_search(small_dataset, jql="project = TA", nextPageToken="garbage")


def test_bad_max_results_is_refused(small_dataset):
    with pytest.raises(ApiError, match="maxResults"):
        run_search(small_dataset, jql="project = TA", maxResults="x")
    with pytest.raises(ApiError, match="at least 1"):
        run_search(small_dataset, jql="project = TA", maxResults=0)


def test_approximate_count(small_dataset):
    assert api.approximate_count(small_dataset, {"jql": 'project in ("TA","TB")'}) == {"count": 4}


# --- bulk changelog ---------------------------------------------------------------------------------

def bulk(ds, **body):
    return api.bulk_changelog(ds, body, tz=TZ, bulk_clamp=body.pop("clamp", 1000))


def test_changelog_has_issue_ids_and_status_items(small_dataset):
    body = bulk(small_dataset, issueIdsOrKeys=["TA-2"])
    log = body["issueChangeLogs"][0]
    assert log["issueId"] == "10002" and body["nextPageToken"] is None
    first = log["changeHistories"][0]
    ids = StatusCatalog.from_params().ids
    assert first["created"] == "2026-10-02T12:00:00.000+0200"
    assert first["author"] == {"accountId": "acc-ann", "displayName": "Ann Martin"}
    assert first["items"] == [{"field": "status", "fieldtype": "jira", "fieldId": "status",
                               "from": ids["Nouveau"], "fromString": "Nouveau",
                               "to": ids["En analyse"], "toString": "En analyse"}]
    assert [h["items"][0]["toString"] for h in log["changeHistories"]] == [
        "En analyse", "Dépriorisé", "Nouveau", "En analyse"]


def test_changelog_accepts_ids_and_keys_and_orders_changes_oldest_first(small_dataset):
    body = bulk(small_dataset, issueIdsOrKeys=["10004", "TA-2"])
    for log in body["issueChangeLogs"]:                      # grouped by issue, oldest change first in each
        times = [h["created"] for h in log["changeHistories"]]
        assert times == sorted(times)
    assert {log["issueId"] for log in body["issueChangeLogs"]} == {"10004", "10002"}


def test_changelog_pages_by_token(small_dataset):
    total, token, pages = 0, None, 0
    while True:
        body = api.bulk_changelog(small_dataset, {"issueIdsOrKeys": ["TA-1", "TA-2", "TB-1"], "maxResults": 4,
                                                  "nextPageToken": token}, tz=TZ, bulk_clamp=1000)
        total += sum(len(log["changeHistories"]) for log in body["issueChangeLogs"])
        pages += 1
        token = body.get("nextPageToken")
        if not token:
            break
    assert (total, pages) == (13, 4)          # 7 + 4 + 2 changes, 4 per page


def test_changelog_refuses_more_than_1000_issues_and_empty_lists(small_dataset):
    with pytest.raises(ApiError, match="at most 1000"):
        bulk(small_dataset, issueIdsOrKeys=[f"X-{n}" for n in range(1001)])
    with pytest.raises(ApiError, match="required"):
        bulk(small_dataset, issueIdsOrKeys=[])


def test_changelog_refuses_more_than_10_field_ids(small_dataset):
    with pytest.raises(ApiError, match="at most 10"):
        bulk(small_dataset, issueIdsOrKeys=["TA-1"], fieldIds=[f"f{n}" for n in range(11)])


def test_changelog_field_filter(small_dataset):
    assert bulk(small_dataset, issueIdsOrKeys=["TA-1"], fieldIds=["summary"])["issueChangeLogs"] == []
    assert len(bulk(small_dataset, issueIdsOrKeys=["TA-1"], fieldIds=["status"])["issueChangeLogs"]) == 1


def test_jira_time_formats_negative_and_zero_offsets():
    moment = dt.datetime(2026, 1, 2, 3, 4, 5, 678000)
    assert api.jira_time(moment, 0) == "2026-01-02T03:04:05.678+0000"
    assert api.jira_time(moment, -330) == "2026-01-01T21:34:05.678-0530"
