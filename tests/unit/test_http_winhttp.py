"""The calls of the Python path through WinHTTP against the fake Jira (no Excel; the real Jira is never called).

The waits between two attempts are recorded, not slept.
"""
import logging
import socket

import pytest

from jira2dashboard.mockjira.generator import generate
from jira2dashboard.mockjira.server import MOCK_EMAIL, MOCK_TOKEN, MockJira
from jira2dashboard.params.model import load_all, project_keys
from jira2dashboard.standalone import jira_api as J
from jira2dashboard.standalone.http_winhttp import JiraClient, WinHttp

PROJECTS = [row for s in load_all() if s.name == "Projects" for row in s.rows]
AUTH = J.basic_auth(MOCK_EMAIL, MOCK_TOKEN)


@pytest.fixture(scope="module")
def dataset():
    return generate(seed=5, issues=60, projects=project_keys())


@pytest.fixture
def jira(dataset):
    with MockJira(dataset, credentials=(MOCK_EMAIL, MOCK_TOKEN), page_clamp=25) as server:
        yield server


def client(base_url, waits, auth=AUTH, retries=4, proxy="", windows_auth=True, timeout_s=5):
    return JiraClient(base_url, auth, retries, WinHttp(timeout_s, proxy, "", windows_auth), sleep=waits.append,
                      rnd=lambda: 0.0)


def test_every_page_of_the_search_and_of_the_changelogs(jira, dataset):
    waits = []
    c = client(jira.base_url, waits)
    issues = c.search_all(J.projects_jql(PROJECTS), J.search_fields("", {}))
    assert sorted(i["key"] for i in issues) == sorted(i.key for i in dataset.issues) and waits == []
    statuses = {i["fields"]["status"]["name"] for i in issues}
    assert "Terminé" in statuses or "À déployer" in statuses               # UTF-8 read from the bytes of the answer
    pages = c.changelogs([i["key"] for i in issues])
    assert sum(len(log["changeHistories"]) for p in pages for log in p["issueChangeLogs"]) == sum(
        len(i.changes) + len(i.key_changes) for i in dataset.issues)
    assert any(r.method == "POST" and r.path == "/rest/api/3/changelog/bulkfetch" for r in jira.requests)


def test_a_wrong_account_stops_at_once_with_the_hint_of_the_vba(jira):
    waits = []
    with pytest.raises(J.JiraError, match=r"HTTP 401 after 1 attempt\(s\).*check Email and ApiToken") as error:
        client(jira.base_url, waits, auth=J.basic_auth(MOCK_EMAIL, "wrong")).call("GET", "/rest/api/3/myself", "me")
    assert error.value.code == 602 and waits == []


def test_server_errors_are_retried_and_each_attempt_is_logged(jira, caplog):
    waits = []
    jira.fail_next("GET", "/rest/api/3/myself", 503, times=2, body={"errorMessages": ["Jira is down"]})
    with caplog.at_level(logging.WARNING, logger="jira2dashboard.extract"):
        assert client(jira.base_url, waits).call("GET", "/rest/api/3/myself", "me")["accountId"] == "acc-mock"
    assert waits == [1.0, 2.0]                                             # 2 s x 2^attempt, half jitter at 0
    lines = [r.getMessage() for r in caplog.records]
    assert len(lines) == 2 and "me: GET /rest/api/3/myself -> HTTP 503 (attempt 1 of 5), retry in 1.0 s. " \
                                "Jira is down" == lines[0]


def test_retry_after_is_never_cut_short(jira):
    waits = []
    jira.fail_next("GET", "/rest/api/3/myself", 429, retry_after="7")
    client(jira.base_url, waits).call("GET", "/rest/api/3/myself", "me")
    assert waits == [7.0]


def test_an_error_that_lasts_ends_after_max_retries_plus_one_attempts(jira, caplog):
    waits = []
    jira.fail_next("GET", "/rest/api/3/myself", 500, times=10, body={"errorMessages": ["boom"]})
    with pytest.raises(J.JiraError, match=r"HTTP 500 after 3 attempt\(s\). boom"):
        client(jira.base_url, waits, retries=2).call("GET", "/rest/api/3/myself", "me")
    assert len(waits) == 2


def free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def test_no_answer_stops_after_three_network_failures():
    waits = []
    with pytest.raises(J.JiraError, match=r"HTTP 0 after 3 attempt\(s\). 0x8007.*Hint: no answer"):
        client(f"http://127.0.0.1:{free_port()}", waits, retries=6, timeout_s=2).call("GET", "/rest/api/3/myself", "me")
    assert len(waits) == 2


def test_a_proxy_that_asks_for_the_windows_session_gets_it_only_with_proxy_windows_auth(dataset):
    """The fake Jira is the proxy, as in the VBA test M_ThroughProxy: Jira is reached under a .invalid name."""
    with MockJira(dataset, credentials=(MOCK_EMAIL, MOCK_TOKEN), proxy_challenge="Negotiate") as proxy:
        address = proxy.base_url.removeprefix("http://")
        answer = client("http://jira.mock.invalid", [], proxy=address).call("GET", "/rest/api/3/myself", "me")
        assert answer["accountId"] == "acc-mock"
        assert any(r.headers.get("proxy-authorization") for r in proxy.requests)    # logged in lower case
        with pytest.raises(J.JiraError, match=r"HTTP 407 .*the proxy asks for authentication"):
            client("http://jira.mock.invalid", [], proxy=address, windows_auth=False).call(
                "GET", "/rest/api/3/myself", "me")
