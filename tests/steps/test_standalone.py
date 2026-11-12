"""Steps of features/standalone.feature: the Python path without VBA, against the fake Jira (in process).

The reference is the flow reference on the fake Jira's data, to the second (Jira's milliseconds are dropped, as the
VBA does). Only the last scenario drives Excel (marker excel).
"""
import base64
import dataclasses
import datetime as dt
import zipfile

import pytest
from pytest_bdd import given, parsers, scenarios, then, when

from jira2dashboard import mock_workbook as M
from jira2dashboard.build_workbook import build_delivery
from jira2dashboard.excel import com
from jira2dashboard.flow import dashboard_spec as D
from jira2dashboard.flow.blueprint import team_names
from jira2dashboard.flow.checks import PER_ISSUE, CheckSettings, Moves, fill_rates, run_checks
from jira2dashboard.flow.itemflow import item_flow
from jira2dashboard.flow.load import daily_load
from jira2dashboard.flow.metrics import BUCKETS, feature_predictability, team_iteration_metrics
from jira2dashboard.flow.model import NOT_COMPUTABLE, Change, Rules
from jira2dashboard.mockjira.anomalies import PI, PI_WINDOW, with_anomalies
from jira2dashboard.mockjira.flow_adapter import to_flow_issues
from jira2dashboard.mockjira.generator import generate
from jira2dashboard.mockjira.server import MOCK_EMAIL, MOCK_TOKEN, MockJira
from jira2dashboard.params.model import load_all, project_keys
from jira2dashboard.paths import INTER_DIR
from jira2dashboard.standalone import __main__ as extract
from jira2dashboard.standalone import results as R
from jira2dashboard.standalone.proxy import Detected
from jira2dashboard.standalone.xlsx_read import read_sheets
from jira2dashboard.standalone.xlsx_write import EPOCH

scenarios("../features/standalone.feature")

TOLERANCE = 1e-9
SEARCH = "/rest/api/3/search/jql"
SOURCE = INTER_DIR / "standalone_source.xlsm"
COPY = INTER_DIR / "standalone_copy.xlsm"


def direct(url):
    """No test depends on the proxy of the machine."""
    return Detected("", "", "direct (test)")


@pytest.fixture
def world(tmp_path):
    return {"tmp": tmp_path, "waits": [], "values": {"ProxyServer": "DIRECT"}, "out": tmp_path / "jira2dashboard.xlsx",
            "log": tmp_path / "python.log"}


@pytest.fixture
def servers():
    started = []
    yield started
    for server in started:
        server.stop()


# --- given -------------------------------------------------------------------------------------------

@given(parsers.parse("un faux Jira avec {n:d} tickets générés et les anomalies de données connues"))
def _jira(world, servers, n):
    world["dataset"], world["anomalies"] = with_anomalies(generate(seed=11, issues=n, projects=project_keys()),
                                                          project_keys())
    servers.append(MockJira(world["dataset"], credentials=(MOCK_EMAIL, MOCK_TOKEN)).start())
    world["jira"] = servers[-1]


@given("un classeur de paramètres rempli pour ce faux Jira")
def _params(world):
    world["url"] = world["jira"].base_url


@given(parsers.parse('le paramètre "{name}" vidé dans le classeur de paramètres'))
def _emptied(world, name):
    world["values"][name] = ""


@given("un fichier de résultats précédent")
def _previous(world):
    world["out"].write_bytes(b"the results of an earlier run")


@given(parsers.parse('le faux Jira répond 500 "{message}" à toutes les recherches'))
def _down(world, message):
    world["jira"].fail_next("GET", SEARCH, 500, times=100, body={"errorMessages": [message], "errors": {}})


@given(parsers.parse("le faux Jira répond 429 avec Retry-After {seconds:d} à la première recherche"))
def _rate_limit(world, seconds):
    world["jira"].fail_next("GET", SEARCH, 429, retry_after=str(seconds))


@given("un proxy devant le faux Jira, typé dans ProxyServer")
def _proxy(world, servers):
    """The fake Jira is the proxy (absolute-form requests) and Jira is reached under a .invalid name. No Windows
    authentication here: the fake proxy accepts the first NTLM leg, which WinHTTP sends with an empty body, so a POST
    would arrive empty (a real proxy answers that leg with a challenge). The Windows session sent to a proxy is tested
    with GET in tests/unit/test_http_winhttp.py, as the VBA test M_ThroughProxy does."""
    servers.append(MockJira(world["dataset"], credentials=(MOCK_EMAIL, MOCK_TOKEN)).start())
    world["proxy"] = servers[-1]
    world["values"] |= {"BaseUrl": "http://jira.mock.invalid",
                        "ProxyServer": world["proxy"].base_url.removeprefix("http://")}


@pytest.fixture(scope="module")
def delivered():
    """The workbook to deliver, built once for the module by the code of `run.ps1 build`."""
    build_delivery(SOURCE)
    return SOURCE


@given("la copie d'essai manuel du classeur à livrer, remplie par Excel pour ce faux Jira")
def _copy(world, delivered, monkeypatch):
    monkeypatch.setattr(M, "COPY", COPY)
    port = world["jira"].base_url.rsplit(":", 1)[1]
    assert M.main(["--workbook", str(delivered), "--port", port]) == 0
    world["workbook"] = COPY


# --- when --------------------------------------------------------------------------------------------

def launch(world, workbook):
    world["code"] = extract.main(["--workbook", str(workbook), "--out", str(world["out"]), "--log", str(world["log"])],
                                 sleep=world["waits"].append, detect=direct)
    world["log_text"] = world["log"].read_text(encoding="utf-8") if world["log"].exists() else ""


@when("je lance l'extraction Python")
def _run(world):
    launch(world, M.write_params_book(world["tmp"] / "params.xlsx", world["url"], world["values"]))


@when("je lance l'extraction Python sur cette copie")
def _run_copy(world):
    launch(world, world["workbook"])


# --- then --------------------------------------------------------------------------------------------

@then(parsers.parse("elle se termine avec le code {code:d}"))
def _code(world, code):
    assert world["code"] == code, world["log_text"][-3000:]


def sheets(world) -> dict[str, list[dict]]:
    if "sheets" not in world:
        grids = read_sheets(world["out"])
        world["sheets"] = {name: [dict(zip(grid[0], row + [None] * (len(grid[0]) - len(row)))) for row in grid[1:]]
                           for name, grid in grids.items() if grid}          # empty cells at the end are not stored
    return world["sheets"]


def to_second(issue):
    def floor(value):
        return value.replace(microsecond=0) if isinstance(value, dt.datetime) else value
    changes = tuple(Change(c.id, floor(c.at), c.from_status, c.to_status) for c in issue.changes)
    return dataclasses.replace(issue, **{f.name: floor(getattr(issue, f.name)) for f in dataclasses.fields(issue)
                                         if isinstance(getattr(issue, f.name), dt.datetime)}, changes=changes)


def reference(world):
    if "reference" not in world:
        dataset = world["dataset"]
        as_of, rules = dataset.as_of, Rules.from_params()
        issues = [to_second(i) for i in to_flow_issues(dataset)]
        settings = dataclasses.replace(CheckSettings.from_params(), pis={PI: PI_WINDOW},
                                       fields={name: getattr(dataset.fields, attr) for name, attr in M.FIELD_IDS.items()})
        items = [item_flow(i, as_of, rules) for i in issues]
        projects = [row for s in load_all() if s.name == "Projects" for row in s.rows]
        world["reference"] = dict(
            as_of=as_of, issues=issues, items=items, teams=team_names(projects), iterations=M.iterations(),
            flags=run_checks(issues, as_of, rules, Moves.from_params(), settings), settings=settings,
            fill=fill_rates(issues, settings, sorted({i.team for i in issues}), sorted({i.type for i in issues})))
    return world["reference"]


def moment(serial):
    return None if serial in (None, "") else EPOCH + dt.timedelta(days=serial)


def same(got, wanted) -> bool:
    if wanted is None or wanted == "":
        return got in (None, "")
    if isinstance(wanted, str) or isinstance(got, str):
        return got == wanted
    if isinstance(wanted, dt.datetime):
        return got is not None and abs(moment(got) - wanted) < dt.timedelta(seconds=0.5)
    return got is not None and abs(got - wanted) <= TOLERANCE * max(1.0, abs(wanted))


@then("la feuille Issues a une ligne par ticket du faux Jira")
def _issues(world):
    assert sorted(r["Key"] for r in sheets(world)["Issues"]) == sorted(i.key for i in world["dataset"].issues)


@then("les feuilles Metrics, PIMetrics, DailyLoad, ItemFlow, FillRates et DataChecks reproduisent la référence Python")
def _results(world):
    ref, got = reference(world), sheets(world)
    as_of, items, wrong = ref["as_of"], ref["items"], []
    metrics = {(r["Iteration-label"], r["Team"]): r for r in got["Metrics"]}
    assert set(metrics) == {(p.name, t) for p in ref["iterations"] for t in ref["teams"]}
    for iteration in ref["iterations"]:
        for team in ref["teams"]:
            m, row = team_iteration_metrics(items, team, iteration, "A", as_of), metrics[(iteration.name, team)]
            pairs = {"Velocity": m.velocity, "FlowTimeCount": m.flow_time_count, "FlowTimeMedian": m.flow_time_median,
                     "FlowTimeP85": m.flow_time_p85, "FlowLoadAverage": m.flow_load_average,
                     "FlowLoadEnd": m.flow_load_end, "FlowEfficiency": m.flow_efficiency,
                     "FlowPredictability": m.flow_predictability}
            for bucket in BUCKETS:
                for column, shares in ((f"DistDone{bucket}", m.distribution_done), (f"DistWip{bucket}", m.distribution_wip)):
                    pairs[column] = NOT_COMPUTABLE if shares == NOT_COMPUTABLE else shares[bucket]
            wrong += [(iteration.name, team, c, row[c], v) for c, v in pairs.items() if not same(row[c], v)]
    first, end = ref["iterations"][0].start, ref["iterations"][-1].end
    loads = {team: daily_load(items, "A", team, first, end, as_of) for team in ref["teams"]}
    for row in got["DailyLoad"]:
        day = moment(row["Date"])
        if not same(row["Load"], loads[row["Team"]].get(day)):
            wrong.append(("DailyLoad", day, row["Team"], row["Load"]))
    assert len(got["DailyLoad"]) == len(R.days_of(first, end)) * len(ref["teams"])
    rows = {r["Key"]: r for r in got["ItemFlow"]}
    for item in items:
        row = rows[item.key]
        expected = {"Team": item.team, "StartUtc": item.start, "DoneUtc": item.done, "FlowTimeDays": item.flow_time_days,
                    "ActiveDays": item.active_days, "WaitDays": item.wait_days, "BacklogReturns": item.backlog_returns,
                    "Excluded": 1 if item.excluded_from_flow_time else 0}
        wrong += [(item.key, c, row[c], v) for c, v in expected.items() if not same(row[c], v)]
        if {c for c in PER_ISSUE if row[c] == 1} != set(ref["flags"].get(item.key, ())):
            wrong.append((item.key, "flags"))
    fill = {(r["Field"], r["Team"], r["Type"]): (r["Count"], r["Filled"], r["Below"]) for r in got["FillRates"]}
    assert fill == {(r.field, r.team, r.type): (r.count, r.filled, 1 if r.below else 0) for r in ref["fill"]}
    expected_checks = [tuple(None if v == "" else v for v in row)
                       for row in R.data_check_rows(ref["flags"], ref["fill"], ref["settings"])]
    assert [tuple(r.values()) for r in got["DataChecks"]] == expected_checks
    pis = feature_predictability(ref["issues"], items, ref["settings"])
    assert [r["PI"] for r in got["PIMetrics"]] == [p.pi for p in pis] and any(p.planned for p in pis)
    for p, row in zip(pis, got["PIMetrics"]):
        wrong += [("PIMetrics", p.pi, c, row[c], v) for c, v in (("FeaturesPlanned", p.planned), ("FeaturesDone", p.done),
                                                                   ("FeaturePredictability", p.ratio)) if not same(row[c], v)]
    assert wrong == [], wrong[:8]
    assert any(r["Velocity"] for r in got["Metrics"]) and any(r["Load"] for r in got["DailyLoad"])


@then("chaque anomalie connue est trouvée par les contrôles qui lui correspondent et seulement par eux")
def _anomalies(world):
    rows = {r["Key"]: r for r in sheets(world)["ItemFlow"]}
    found = {key: {c for c in PER_ISSUE if rows[key][c] == 1} for key in world["anomalies"]}
    assert found == {key: set(ids) for key, ids in world["anomalies"].items()}


@then("la feuille Charts a les 8 graphiques du Dashboard en SVG")
def _charts(world):
    with zipfile.ZipFile(world["out"]) as package:
        names = package.namelist()
        drawing = package.read("xl/drawings/drawing1.xml").decode("utf-8")
    assert sum(1 for n in names if n.endswith(".svg")) == len(D.CHARTS)
    assert all(f'name="{spec.name}"' in drawing for spec in D.CHARTS)
    assert [s["Item"] for s in sheets(world)["Run"]][:1] == ["Made by"]


@then("la feuille Config des résultats masque le token")
def _masked(world):
    config = {r["Parameter"]: r["Value"] for r in sheets(world)["Config"]}
    assert config["ApiToken"] == extract.MASK and config["Email"] == MOCK_EMAIL
    with zipfile.ZipFile(world["out"]) as package:
        assert not any(MOCK_TOKEN.encode() in package.read(name) for name in package.namelist())


@then("le faux Jira n'a reçu aucun appel")
def _no_call(world):
    assert world["jira"].requests == []


@then(parsers.parse('le journal dit "{text}"'))
def _log_says(world, text):
    assert text in world["log_text"], world["log_text"][-2000:]


@then(parsers.parse('le journal donne pour chaque tentative le statut 500, l\'appel search et le message "{message}"'))
def _each_attempt(world, message):
    lines = [line for line in world["log_text"].splitlines() if f"search: GET {SEARCH} -> HTTP 500 (attempt" in line]
    assert len(lines) == 5 and all(line.endswith(message) for line in lines), lines
    assert len(world["waits"]) == 4


@then("le journal ne contient ni le token ni l'en-tête Authorization")
def _no_secret(world):
    basic = base64.b64encode(f"{MOCK_EMAIL}:{MOCK_TOKEN}".encode()).decode()
    assert MOCK_TOKEN not in world["log_text"] and basic not in world["log_text"]


@then("le fichier de résultats précédent n'a pas changé")
def _unchanged(world):
    assert world["out"].read_bytes() == b"the results of an earlier run"


@then(parsers.parse("l'extraction a attendu {seconds:d} secondes avant de recommencer"))
def _waited(world, seconds):
    assert len(world["waits"]) == 1 and seconds <= world["waits"][0] < seconds + 0.5


@then("la feuille Errors garde l'avertissement du 429")
def _warning(world):
    messages = [r["Message"] for r in sheets(world)["Errors"]]
    assert any("HTTP 429" in m and "Retry-After 7" in m for m in messages), messages


@then("la recherche et les historiques sont passés par le proxy")
def _through_proxy(world):
    paths = {r.path for r in world["proxy"].requests}
    assert {SEARCH, "/rest/api/3/changelog/bulkfetch"} <= paths and world["jira"].requests == []


@then("Excel ouvre le fichier de résultats et ses 8 graphiques sont des images SVG")
def _excel_opens(world):
    with com.excel_session() as app:
        app.DisplayAlerts = False
        wb = app.Workbooks.Open(str(world["out"].resolve()))
        charts = wb.Worksheets("Charts")
        kinds = [charts.Shapes(n).Type for n in range(1, charts.Shapes.Count + 1)]
        names = [wb.Worksheets(n).Name for n in range(1, wb.Worksheets.Count + 1)]
        wb.Close(False)
        del charts, wb, app
    assert kinds == [com.MSO_GRAPHIC] * len(D.CHARTS)
    assert names[:3] == ["Run", "Charts", "Metrics"] and "Errors" in names
