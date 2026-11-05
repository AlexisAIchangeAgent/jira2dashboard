"""The rules of the calls to Jira of the Python path (pure): the same as the VBA modules modJira and modHttp."""
import base64
import json

import pytest

from jira2dashboard.mockjira import api, jql
from jira2dashboard.mockjira.generator import generate
from jira2dashboard.params.model import load_all, project_keys
from jira2dashboard.standalone import jira_api as J

PROJECTS = [("TA", "Team", "Alpha"), ("PRG", "Program", ""), ("TB", "Team", "Beta")]


def test_the_jql_takes_the_program_first_then_the_filter_and_a_fixed_order():
    assert J.projects_jql(PROJECTS) == 'project in ("PRG","TA","TB") ORDER BY key ASC'
    assert J.projects_jql(PROJECTS, " updated >= -90d ") == (
        'project in ("PRG","TA","TB") AND (updated >= -90d) ORDER BY key ASC')
    assert J.projects_jql([(' A"B ', "Team", "x")]) == 'project in ("AB") ORDER BY key ASC'


@pytest.mark.parametrize("projects, extra, code", [([], "", 600), (PROJECTS, "status = Done order by key", 603)])
def test_the_jql_is_refused_like_in_the_vba(projects, extra, code):
    with pytest.raises(J.JiraError) as error:
        J.projects_jql(projects, extra)
    assert error.value.code == code


def test_the_fake_jira_understands_the_jql():
    projects = [row for s in load_all() if s.name == "Projects" for row in s.rows]
    dataset = generate(issues=40, projects=project_keys())
    page = api.search(dataset, {"jql": J.projects_jql(projects), "fields": ["key"], "maxResults": 100},
                      tz=120, base_url="http://x", page_clamp=100)
    assert len(page["issues"]) == 40 and page["isLast"] is True


def test_the_fields_asked_for_are_the_system_ones_then_the_configured_ids():
    fields = {"StoryPointsField": "cf_sp", "BusinessValueField": "", "PIField": "cf_pi", "TeamField": "",
              "SprintField": "cf_sprint", "TargetStartField": "", "TargetEndField": "cf_end"}
    assert J.search_fields("cf_wt", fields) == list(J.ALWAYS_FIELDS) + ["cf_wt", "cf_sp", "cf_pi", "cf_sprint", "cf_end"]
    assert J.search_fields("", {}) == list(J.ALWAYS_FIELDS)


def test_the_search_path_encodes_like_the_vba():
    path = J.search_path('project in ("PRG") ORDER BY key ASC', ["key", "status"], "a+b/c=")
    assert path == ("/rest/api/3/search/jql?jql=project%20in%20%28%22PRG%22%29%20ORDER%20BY%20key%20ASC"
                    "&maxResults=100&fields=key%2Cstatus&nextPageToken=a%2Bb%2Fc%3D")
    assert J.encode("é-._~") == "%C3%A9-._~"


def test_the_changelog_body_and_its_chunks():
    assert json.loads(J.changelog_body(["A-1", "A-2"])) == {"issueIdsOrKeys": ["A-1", "A-2"], "maxResults": 1000}
    assert json.loads(J.changelog_body(["A-1"], "tok"))["nextPageToken"] == "tok"
    assert [len(c) for c in J.chunks([str(n) for n in range(2500)])] == [1000, 1000, 500]


@pytest.mark.parametrize("value, url", [(" https://site.atlassian.net// ", "https://site.atlassian.net"),
                                        ("http://127.0.0.1:8099/", "http://127.0.0.1:8099"),
                                        ("http://jira.invalid", "http://jira.invalid"), ("", "")])
def test_the_base_url_is_https_except_for_the_fake_jira(value, url):
    assert J.check_base_url(value) == url


@pytest.mark.parametrize("value", ["http://site.atlassian.net", "http://localhost.evil.example", "ftp://x"])
def test_a_base_url_in_clear_text_is_refused(value):
    with pytest.raises(J.JiraError, match="BaseUrl must start with https://"):
        J.check_base_url(value)


def test_the_proxy_loses_its_scheme_and_the_account_is_basic():
    assert J.normalize_proxy(" HTTP://proxy.name.com:8080/ ") == "proxy.name.com:8080"
    assert J.normalize_proxy("https://p:1") == "p:1" and J.normalize_proxy("") == ""
    header = J.basic_auth(" me@x.com\t", " tok\r\n")
    assert base64.b64decode(header.removeprefix("Basic ")) == b"me@x.com:tok"


def test_retries_follow_429_and_the_server_errors_only():
    assert [s for s in (200, 400, 401, 404, 407, 429, 500, 503, 599) if J.is_retryable(s)] == [429, 500, 503, 599]


def test_the_wait_grows_and_never_beats_retry_after():
    assert J.retry_delay(0, None, 0.0) == 1.0 and J.retry_delay(0, None, 0.999) == pytest.approx(1.999)
    assert J.retry_delay(3, "", 0.0) == 8.0 and J.retry_delay(10, None, 0.0) == 15.0          # capped at 30 s
    assert J.retry_delay(0, "7", 0.0) == 7.0 and J.retry_delay(0, "500", 0.0) == 120.0
    assert J.retry_delay(0, "Wed, 21 Oct 2026 07:28:00 GMT", 0.0) == 1.0                   # a date: not understood


def test_an_error_names_the_status_the_attempts_what_jira_said_and_a_hint():
    error = J.status_error(401, 1, "", json.dumps({"errorMessages": ["Bad credentials"], "errors": {"jql": "x"}}))
    assert error.code == 602
    assert str(error) == ("HTTP 401 after 1 attempt(s). Bad credentials; jql: x Hint: Jira refused the account. "
                          "In the Config sheet check Email and ApiToken.")
    assert "Hint: no answer" in str(J.status_error(0, 3, "-2147012867 cannot connect", ""))
    assert str(J.status_error(500, 5, "", "<html>oops</html>")).startswith("HTTP 500 after 5 attempt(s). <html>oops")


def test_an_answer_that_is_not_json_is_named():
    assert J.parse_json(' {"a": 1}', "search") == {"a": 1}
    with pytest.raises(J.JiraError, match="not JSON") as error:
        J.parse_json("<html>login</html>", "search")
    assert error.value.code == 607
