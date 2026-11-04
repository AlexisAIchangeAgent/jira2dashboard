"""The fake Jira over real HTTP (loopback)."""
import base64
import json
import urllib.parse

import pytest

from jira2dashboard.mockjira.server import MockJira

SEARCH = "/rest/api/3/search/jql"


@pytest.fixture
def jira(small_dataset):
    with MockJira(small_dataset) as server:
        yield server


def get_search_path(jql, **extra):
    """The URL shape built by modJira.JiraSearchUrl."""
    params = {"jql": jql, "maxResults": "100", "fields": "*all", "expand": "names", **extra}
    return SEARCH + "?" + urllib.parse.urlencode(params, quote_via=urllib.parse.quote)


def test_get_search_like_the_vba_client(jira, call):
    result = call(jira, "GET", get_search_path('project in ("TA","TB") ORDER BY key ASC'))
    body = result.json
    assert result.status == 200 and "application/json" in result.headers["Content-Type"]
    assert [i["key"] for i in body["issues"]] == ["TA-1", "TA-2", "TA-3", "TB-1"]
    assert body["isLast"] is True and body["names"]["summary"] == "Summary"


def test_post_search_with_a_json_body(jira, call):
    result = call(jira, "POST", SEARCH, {"jql": "project = TA", "fields": ["summary"], "maxResults": 2})
    assert [i["key"] for i in result.json["issues"]] == ["TA-1", "TA-2"]
    assert "nextPageToken" in result.json


def test_utf8_is_sent_as_utf8(jira, call):
    raw = call(jira, "GET", get_search_path("project = TA")).raw
    assert "Terminé".encode("utf-8") in raw


def test_approximate_count_and_field_catalog_and_myself(jira, call):
    assert call(jira, "POST", "/rest/api/3/search/approximate-count", {"jql": "project = TA"}).json == {"count": 3}
    assert any(f["id"] == "status" for f in call(jira, "GET", "/rest/api/3/field").json)
    assert call(jira, "GET", "/rest/api/3/myself").json["accountId"] == "acc-mock"


def test_bulk_changelog_over_http(jira, call):
    result = call(jira, "POST", "/rest/api/3/changelog/bulkfetch", {"issueIdsOrKeys": ["TA-1"], "maxResults": 1000})
    assert result.status == 200 and len(result.json["issueChangeLogs"][0]["changeHistories"]) == 7


def test_the_removed_search_endpoints_answer_410(jira, call):
    for path in ("/rest/api/3/search", "/rest/api/3/search/id"):
        result = call(jira, "GET", path + "?jql=project%3DTA")
        assert result.status == 410 and "/search/jql" in result.json["errorMessages"][0]


def test_unbounded_jql_answers_400(jira, call):
    result = call(jira, "GET", get_search_path("ORDER BY key ASC"))
    assert result.status == 400 and "Unbounded" in result.json["errorMessages"][0]


def test_unknown_path_is_404_and_wrong_method_is_405(jira, call):
    assert call(jira, "GET", "/rest/api/3/nothing").status == 404
    assert call(jira, "GET", "/rest/api/3/changelog/bulkfetch").status == 405


def test_credentials_are_checked_when_configured(small_dataset, call):
    good = "Basic " + base64.b64encode(b"me@example.com:tok").decode()
    with MockJira(small_dataset, credentials=("me@example.com", "tok")) as jira:
        assert call(jira, "GET", "/rest/api/3/myself").status == 401
        assert call(jira, "GET", "/rest/api/3/myself").raw == b"Unauthorized"       # plain text, like Jira
        assert call(jira, "GET", "/rest/api/3/myself", headers={"Authorization": "Basic eA=="}).status == 401
        assert call(jira, "GET", "/rest/api/3/myself", headers={"Authorization": good}).status == 200


def test_credentials_must_match_exactly_like_the_powershell_script_builds_them(small_dataset, call):
    """Authorization = Basic base64("email:token") with the token trimmed by the client."""
    email, token = "genericemail@domain.com", "tok-123"
    header = "Basic " + base64.b64encode(f"{email}:{token}".encode("ascii")).decode()
    with MockJira(small_dataset, credentials=(email, token)) as jira:
        assert call(jira, "GET", "/rest/api/3/myself", headers={"Authorization": header}).status == 200
        padded = "Basic " + base64.b64encode(f"{email}:{token} \r\n".encode("ascii")).decode()
        assert call(jira, "GET", "/rest/api/3/myself", headers={"Authorization": padded}).status == 401


def test_the_server_answers_when_used_as_an_http_proxy(small_dataset):
    """Through WinHTTP SetProxy the client sends the absolute form `GET http://<site>/rest/...` to the proxy."""
    import urllib.request
    with MockJira(small_dataset) as jira:
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({"http": jira.base_url}))
        with opener.open("http://jira.mock.invalid/rest/api/3/myself", timeout=10) as response:
            assert response.status == 200 and json.loads(response.read())["accountId"] == "acc-mock"
        logged = jira.requests[-1]
        assert logged.path == "/rest/api/3/myself" and logged.headers["host"] == "jira.mock.invalid"


def _via_proxy(jira, url, headers=None):
    import urllib.error
    import urllib.request
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({"http": jira.base_url}))
    request = urllib.request.Request(url, headers=headers or {})
    try:
        with opener.open(request, timeout=10) as response:
            return response.status, dict(response.headers)
    except urllib.error.HTTPError as err:
        return err.code, dict(err.headers)


def test_a_proxy_that_wants_authentication_answers_407_with_its_challenge(small_dataset):
    with MockJira(small_dataset, proxy_challenge="Negotiate") as jira:
        status, headers = _via_proxy(jira, "http://jira.mock.invalid/rest/api/3/myself")
        assert status == 407 and headers["Proxy-Authenticate"] == "Negotiate"


def test_the_request_goes_through_once_it_carries_proxy_authorization(small_dataset):
    with MockJira(small_dataset, proxy_challenge="Negotiate") as jira:
        status, _ = _via_proxy(jira, "http://jira.mock.invalid/rest/api/3/myself",
                               {"Proxy-Authorization": "Negotiate YIIB"})
        assert status == 200
        assert jira.requests[-1].headers["proxy-authorization"] == "Negotiate YIIB"


def test_direct_requests_are_not_asked_for_proxy_authentication(small_dataset, call):
    with MockJira(small_dataset, proxy_challenge="Negotiate") as jira:
        assert call(jira, "GET", "/rest/api/3/myself").status == 200


def test_connect_opens_a_tunnel_to_the_fake(small_dataset):
    """HTTPS through a proxy: CONNECT host:port, then the exchange runs inside the tunnel (plain HTTP here)."""
    import http.client
    with MockJira(small_dataset) as jira:
        host, port = jira._server.server_address[:2]
        connection = http.client.HTTPConnection(host, port, timeout=10)
        connection.set_tunnel("jira.mock.invalid", 443)
        connection.request("GET", "/rest/api/3/myself")
        response = connection.getresponse()
        assert response.status == 200 and json.loads(response.read())["accountId"] == "acc-mock"
        connection.close()
        assert [r.path for r in jira.requests if r.method == "CONNECT"] == ["jira.mock.invalid:443"]


def test_connect_without_proxy_authorization_is_refused_when_the_proxy_wants_it(small_dataset):
    import http.client
    with MockJira(small_dataset, proxy_challenge="NTLM") as jira:
        host, port = jira._server.server_address[:2]
        refused = http.client.HTTPConnection(host, port, timeout=10)
        refused.set_tunnel("jira.mock.invalid", 443)
        with pytest.raises(OSError, match="407"):
            refused.request("GET", "/rest/api/3/myself")
        refused.close()
        allowed = http.client.HTTPConnection(host, port, timeout=10)
        allowed.set_tunnel("jira.mock.invalid", 443, headers={"Proxy-Authorization": "NTLM TlRMTVNT"})
        allowed.request("GET", "/rest/api/3/myself")
        assert allowed.getresponse().status == 200
        allowed.close()


def test_fail_next_serves_the_fault_then_recovers(jira, call):
    jira.fail_next("GET", SEARCH, 429, times=2, retry_after="3")
    path = get_search_path("project = TA")
    first, second, third = (call(jira, "GET", path) for _ in range(3))
    assert (first.status, second.status, third.status) == (429, 429, 200)
    assert first.headers["Retry-After"] == "3" and first.headers["RateLimit-Reason"] == "jira-burst-based"


def test_503_fault_with_and_without_retry_after(jira, call):
    jira.fail_next("POST", "/rest/api/3/changelog/bulkfetch", 503, retry_after="1")
    jira.fail_next("POST", "/rest/api/3/changelog/bulkfetch", 503)
    body = {"issueIdsOrKeys": ["TA-1"]}
    first = call(jira, "POST", "/rest/api/3/changelog/bulkfetch", body)
    second = call(jira, "POST", "/rest/api/3/changelog/bulkfetch", body)
    assert first.headers.get("Retry-After") == "1" and "Retry-After" not in second.headers
    assert call(jira, "POST", "/rest/api/3/changelog/bulkfetch", body).status == 200


def test_a_fault_hits_only_its_own_method_and_path(jira, call):
    jira.fail_next("GET", SEARCH, 429)
    assert call(jira, "GET", "/rest/api/3/field").status == 200
    assert call(jira, "POST", SEARCH, {"jql": "project = TA"}).status == 200
    assert call(jira, "GET", get_search_path("project = TA")).status == 429


def test_faults_can_be_set_and_reset_over_http_for_the_vba_tests(jira, call):
    call(jira, "POST", "/__mock__/fail", {"method": "GET", "path": "/rest/api/3/field", "status": 429, "times": 1,
                                          "retryAfter": "0"})
    assert call(jira, "GET", "/rest/api/3/field").status == 429
    call(jira, "POST", "/__mock__/fail", {"method": "GET", "path": "/rest/api/3/field", "status": 500})
    call(jira, "POST", "/__mock__/reset")
    assert call(jira, "GET", "/rest/api/3/field").status == 200
    assert len(jira.requests) == 1                                     # reset also cleared the log


def test_requests_are_logged_with_their_body_and_headers(jira, call):
    call(jira, "POST", SEARCH, {"jql": "project = TA"}, headers={"Authorization": "Basic eA=="})
    logged = jira.requests[-1]
    assert (logged.method, logged.path) == ("POST", SEARCH)
    assert json.loads(logged.body) == {"jql": "project = TA"} and logged.headers["authorization"] == "Basic eA=="


def test_near_limit_header_can_be_switched_on(small_dataset, call):
    with MockJira(small_dataset, near_limit=True) as jira:
        assert call(jira, "GET", "/rest/api/3/myself").headers["X-RateLimit-NearLimit"] == "true"


def test_bad_input_is_a_400_with_the_jira_error_envelope(jira, call):
    result = call(jira, "POST", SEARCH, {"jql": "project = TA", "maxResults": "abc"})
    assert result.status == 400 and "maxResults" in result.json["errorMessages"][0]
    broken = call(jira, "POST", SEARCH, None)           # no body: jql missing
    assert broken.status == 400
