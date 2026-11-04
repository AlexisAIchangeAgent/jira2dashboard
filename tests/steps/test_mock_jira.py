"""Steps of features/mock_jira.feature (no Excel): the fake Jira over real HTTP."""
import urllib.parse
from types import SimpleNamespace

import pytest
from pytest_bdd import given, parsers, scenarios, then, when

from jira2dashboard.mockjira.generator import generate
from jira2dashboard.mockjira.server import MockJira
from jira2dashboard.params.model import project_keys

scenarios("../features/mock_jira.feature")
SEARCH = "/rest/api/3/search/jql"


@pytest.fixture
def world():
    return {}


@pytest.fixture
def start_jira():
    """start_jira(dataset, **options) -> running MockJira, stopped at the end of the scenario."""
    servers = []

    def start(dataset, **options):
        server = MockJira(dataset, **options).start()
        servers.append(server)
        return server
    yield start
    for server in servers:
        server.stop()


def search_path(jql, size, token=None):
    """The URL shape built by modJira.JiraSearchUrl."""
    params = {"jql": jql, "maxResults": str(size), "fields": "*all", "expand": "names"}
    if token:
        params["nextPageToken"] = token
    return SEARCH + "?" + urllib.parse.urlencode(params, quote_via=urllib.parse.quote)


# --- given -------------------------------------------------------------------------------------------

@given(parsers.parse("un faux Jira avec {n:d} tickets générés"))
def _generated(world, start_jira, n):
    world["jira"] = start_jira(generate(seed=11, issues=n))


@given(parsers.parse("un faux Jira qui sert au plus {clamp:d} tickets par page avec {n:d} tickets générés"))
def _clamped(world, start_jira, clamp, n):
    world["jira"] = start_jira(generate(seed=11, issues=n), page_clamp=clamp)


@given(parsers.parse('un faux Jira avec {n:d} tickets générés qui exige le compte "{email}" et le token "{token}"'))
def _generated_with_account(world, start_jira, n, email, token):
    world["jira"] = start_jira(generate(seed=11, issues=n), credentials=(email, token))
    world["account"] = (email, token)


@given(parsers.parse('un faux Jira avec {n:d} tickets générés qui exige le compte "{email}" et le token "{token}" '
                     'derrière un proxy qui exige "{scheme}"'))
def _generated_with_account_and_proxy(world, start_jira, n, email, token, scheme):
    """The projects are those of params/projects.md (the sheet Projects of the workbook), the Program first."""
    dataset = generate(seed=11, issues=n, projects=project_keys())
    world["jira"] = start_jira(dataset, credentials=(email, token), proxy_challenge=scheme)
    world["account"] = (email, token)


@given(parsers.parse('un faux Jira avec les cinq tickets écrits à la main derrière un proxy qui exige "{scheme}"'))
def _by_hand_behind_proxy(world, start_jira, small_dataset, scheme):
    world["jira"] = start_jira(small_dataset, proxy_challenge=scheme)


@given("un faux Jira avec les cinq tickets écrits à la main")
def _by_hand(world, start_jira, small_dataset):
    world["jira"] = start_jira(small_dataset)


@given(parsers.parse('un faux Jira avec les cinq tickets écrits à la main qui exige le compte "{email}" et le token "{token}"'))
def _by_hand_with_account(world, start_jira, small_dataset, email, token):
    world["jira"] = start_jira(small_dataset, credentials=(email, token))


@given(parsers.parse('la recherche répond 429 avec Retry-After "{seconds}" deux fois'))
def _rate_limited(world, seconds):
    world["jira"].fail_next("GET", SEARCH, 429, times=2, retry_after=seconds)


# --- when --------------------------------------------------------------------------------------------

@when(parsers.parse('je lis toutes les pages de la recherche de "{projects}" avec {size:d} par page'))
def _read_all_pages(world, call, projects, size):
    jql = "project in (" + ",".join(f'"{p}"' for p in projects.split(",")) + ") ORDER BY key ASC"
    pages, token = [], None
    for _ in range(100):
        result = call(world["jira"], "GET", search_path(jql, size, token))
        assert result.status == 200, result.raw
        pages.append(result.json)
        token = result.json.get("nextPageToken")
        if result.json["isLast"] or not token:
            break
    world["pages"] = pages
    world["issues"] = [issue for page in pages for issue in page["issues"]]


@when("je lis les historiques de tous les tickets reçus")
def _read_histories(world, call):
    ids = [issue["id"] for issue in world["issues"]]
    logs, token = [], None
    for _ in range(100):
        body = {"issueIdsOrKeys": ids[:1000], "maxResults": 1000}
        if token:
            body["nextPageToken"] = token
        result = call(world["jira"], "POST", "/rest/api/3/changelog/bulkfetch", body)
        assert result.status == 200, result.raw
        logs += result.json["issueChangeLogs"]
        token = result.json.get("nextPageToken")
        if not token:
            break
    world["histories"] = {}
    for log in logs:
        world["histories"].setdefault(log["issueId"], []).extend(log["changeHistories"])


@when(parsers.parse('je lis les historiques du ticket "{key}"'))
def _read_one_history(world, call, key):
    result = call(world["jira"], "POST", "/rest/api/3/changelog/bulkfetch", {"issueIdsOrKeys": [key]})
    world["statuses"] = [h["items"][0]["toString"] for log in result.json["issueChangeLogs"]
                         for h in log["changeHistories"]]


@when("je lance trois recherches")
def _three_searches(world, call):
    path = search_path("project = TA", 10)
    world["results"] = [call(world["jira"], "GET", path) for _ in range(3)]


@when(parsers.parse('j\'appelle "{method} {path}"'))
def _call(world, call, method, path):
    world["result"] = call(world["jira"], method, path)


@when(parsers.parse('j\'appelle "{method} {path}" avec le compte "{email}" et le token "{token}"'))
def _call_with_account(world, call, method, path, email, token):
    """Authorization is built like the PowerShell script does: Basic base64("email:token")."""
    import base64
    header = "Basic " + base64.b64encode(f"{email}:{token}".encode("ascii")).decode("ascii")
    world["result"] = call(world["jira"], method, path, headers={"Authorization": header})


@given("le classeur de test ouvert sur ce faux Jira")
def _workbook_on_the_mock(world, built_workbook):
    dataset = world["jira"]._state.dataset
    world["wb"] = built_workbook["wb"]
    changes = sum(len(i.changes) for i in dataset.issues)
    email, token = world["account"]
    _run_vba(world, "MockSetup", world["jira"].base_url, email, token, len(dataset.issues), changes)


def _run_vba(world, macro, *args):
    wb = world["wb"]
    return wb.Application.Run(f"'{wb.Name}'!modTests.{macro}", *args)


@when("le classeur lance ses tests contre le faux Jira")
def _run_mock_tests(world):
    world["summary"] = _run_vba(world, "RunMockTests")


@then("tous les tests VBA contre le faux Jira passent")
def _mock_tests_pass(world):
    rows = world["wb"].Worksheets("TestResults").UsedRange.Value
    mock_rows = [r for r in rows[1:] if r[1] == "mock" or str(r[0]).endswith(".ERR")]
    failures = [list(r[:7]) for r in mock_rows if r[6] != "OK"]
    assert failures == [], failures
    assert len(mock_rows) >= 20, world["summary"]
    assert world["summary"].endswith("0 KO"), world["summary"]


@when(parsers.parse('je lance la recherche "{jql}"'))
def _search(world, call, jql):
    world["result"] = call(world["jira"], "GET", search_path(jql, 10))


# --- then --------------------------------------------------------------------------------------------

@then(parsers.parse("j'ai reçu {n:d} tickets en {pages:d} pages sans doublon"))
def _received(world, n, pages):
    keys = [issue["key"] for issue in world["issues"]]
    assert len(keys) == n and len(set(keys)) == n
    assert len(world["pages"]) == pages


@then("le statut actuel de chaque ticket est la dernière destination de son historique")
def _status_is_last_destination(world):
    checked = 0
    for issue in world["issues"]:
        history = sorted(world["histories"].get(issue["id"], []), key=lambda h: h["created"])
        expected = history[-1]["items"][0]["toString"] if history else "Nouveau"
        assert issue["fields"]["status"]["name"] == expected, issue["key"]
        checked += 1
    assert checked == len(world["issues"]) > 0


@then(parsers.parse('les statuts successifs sont "{statuses}"'))
def _statuses(world, statuses):
    assert world["statuses"] == statuses.split(", ")


@then(parsers.parse("les réponses sont {a:d}, {b:d}, {c:d}"))
def _statuses_of_three(world, a, b, c):
    assert [r.status for r in world["results"]] == [a, b, c]


@then(parsers.parse('la première réponse porte Retry-After "{seconds}"'))
def _retry_after(world, seconds):
    assert world["results"][0].headers["Retry-After"] == seconds


def _through_proxy(jira, url, headers=None):
    import urllib.error
    import urllib.request
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({"http": jira.base_url}))
    try:
        with opener.open(urllib.request.Request(url, headers=headers or {}), timeout=10) as response:
            return response.status, dict(response.headers)
    except urllib.error.HTTPError as err:
        return err.code, dict(err.headers)


@when(parsers.parse('j\'appelle "{method} {url}" par le proxy'))
def _call_via_proxy(world, method, url):
    status, headers = _through_proxy(world["jira"], url)
    world["result"] = SimpleNamespace(status=status, headers=headers)


@when(parsers.parse('j\'appelle "{method} {url}" par le proxy avec "{authorization}"'))
def _call_via_proxy_authorized(world, method, url, authorization):
    status, headers = _through_proxy(world["jira"], url, {"Proxy-Authorization": authorization})
    world["result"] = SimpleNamespace(status=status, headers=headers)


@then(parsers.parse('l\'en-tête de réponse "{name}" vaut "{value}"'))
def _response_header(world, name, value):
    assert world["result"].headers[name] == value


def _open_tunnel(world, target, headers):
    import http.client
    host, port = world["jira"]._server.server_address[:2]
    connection = http.client.HTTPConnection(host, port, timeout=10)
    host_name, _, target_port = target.partition(":")
    connection.set_tunnel(host_name, int(target_port), headers=headers)
    world["connection"] = connection
    try:
        connection.request("GET", "/rest/api/3/myself")
        world["tunnel_status"] = connection.getresponse().status
    except OSError as err:
        world["tunnel_status"] = 407 if "407" in str(err) else None       # http.client: "Tunnel connection failed: 407 ..."
    finally:
        connection.close()


@when(parsers.parse('j\'ouvre un tunnel CONNECT vers "{target}" sans authentification'))
def _tunnel_without_auth(world, target):
    _open_tunnel(world, target, {})


@when(parsers.parse('j\'ouvre un tunnel CONNECT vers "{target}" avec "{authorization}"'))
def _tunnel_with_auth(world, target, authorization):
    _open_tunnel(world, target, {"Proxy-Authorization": authorization})


@then(parsers.parse("le tunnel est refusé avec {status:d}"))
def _tunnel_refused(world, status):
    assert world["tunnel_status"] == status


@then(parsers.parse("la requête dans le tunnel répond {status:d}"))
def _tunnel_answers(world, status):
    assert world["tunnel_status"] == status


@then("le proxy a reçu l'authentification Windows du client")
def _proxy_got_windows_auth(world):
    through = [r for r in world["jira"].requests if r.headers.get("host") == "jira.mock.invalid"]
    with_auth = [r for r in through if r.headers.get("proxy-authorization", "").split(" ")[0] in ("Negotiate", "NTLM")]
    without = [r for r in through if "proxy-authorization" not in r.headers]
    assert with_auth, "ProxyWindowsAuth = Yes: the client must answer the 407 with the Windows session"
    assert without, "ProxyWindowsAuth = No: the client must not send it (the 407 test)"


@then(parsers.parse('le faux Jira a reçu des requêtes adressées au site "{host}"'))
def _requests_for_the_site(world, host):
    hosts = [r.headers.get("host") for r in world["jira"].requests]
    assert host in hosts, set(hosts)


@then(parsers.parse("la réponse est {status:d}"))
def _response_is(world, status):
    assert world["result"].status == status
