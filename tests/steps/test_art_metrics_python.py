"""Steps of features/art_metrics.feature: the Python path (no Excel) on the scenario art-metrics of the fake Jira.

The expected numbers come from tests/art_oracle.py, worked out by hand; nothing is computed with flow/ here.
"""
import pytest
from art_oracle import (ART_ROWS, CHAIN, DONE_A, DONE_B, FLOW_EFFICIENCY, FLOW_TIME_DAYS, ITERATION_STARTS, LONG_STATUS, PI_OF_FEATURE,
                        PIS, PLANNED, TEAM_ROWS, VELOCITY)
from pytest_bdd import given, parsers, scenarios, then, when

from jira2dashboard import mock_workbook as M
from jira2dashboard import sheet_names as S
from jira2dashboard.mockjira.scenarios import get
from jira2dashboard.mockjira.server import MOCK_EMAIL, MOCK_TOKEN, MockJira
from jira2dashboard.standalone import __main__ as extract
from jira2dashboard.standalone.proxy import Detected
from jira2dashboard.standalone.xlsx_read import read_sheets

scenarios("../features/art_metrics.feature")

TOLERANCE = 1e-9


@pytest.fixture
def world(tmp_path):
    return {"tmp": tmp_path, "values": {"ProxyServer": "DIRECT"}, "out": tmp_path / "jira2dashboard.xlsx", "log": tmp_path / "python.log"}


@pytest.fixture
def servers():
    started = []
    yield started
    for server in started:
        server.stop()


@given("le faux Jira du scénario art-metrics")
def _jira(world, servers):
    world["scenario"] = get("art-metrics")
    servers.append(MockJira(world["scenario"].dataset, credentials=(MOCK_EMAIL, MOCK_TOKEN)).start())
    world["jira"] = servers[-1]


@given(parsers.parse('le paramètre "{name}" vaut "{value}" dans le classeur de paramètres'))
def _value(world, name, value):
    world["values"][name] = value


@when("je lance l'extraction Python sur ce scénario")
def _run(world):
    book = M.write_params_book(world["tmp"] / "params.xlsx", world["jira"].base_url, world["values"], world["scenario"])
    world["code"] = extract.main(["--workbook", str(book), "--out", str(world["out"]), "--log", str(world["log"])],
                                 sleep=lambda seconds: None, detect=lambda url: Detected("", "", "direct (test)"))
    world["log_text"] = world["log"].read_text(encoding="utf-8") if world["log"].exists() else ""


@then(parsers.parse("elle se termine avec le code {code:d}"))
def _code(world, code):
    assert world["code"] == code, world["log_text"][-3000:]


def sheet(world, name: str) -> list[dict]:
    """The rows of a result sheet, by column name."""
    grids = read_sheets(world["out"])
    grid = grids[S.of(name)]
    return [dict(zip(grid[0], row + [None] * (len(grid[0]) - len(row)))) for row in grid[1:]]


def close(value, expected) -> bool:
    return value is not None and abs(float(value) - expected) < TOLERANCE


@then("la feuille PIMetrics donne pour les 4 PI 6 prévues et 6, 5, 3 et 4 faites")
def _pimetrics_a(world):
    rows = sheet(world, "PIMetrics")
    assert [r["PI"] for r in rows] == list(PIS)
    assert [(r["FeaturesPlanned"], r["FeaturesDone"]) for r in rows] == [(PLANNED[pi], DONE_A[pi]) for pi in PIS]


@then("la feuille PIMetrics donne pour les 4 PI 6 prévues et 6, 5, 3 et 5 faites")
def _pimetrics_b(world):
    rows = sheet(world, "PIMetrics")
    assert [(r["FeaturesPlanned"], r["FeaturesDone"]) for r in rows] == [(PLANNED[pi], DONE_B[pi]) for pi in PIS]


@then("la Feature Predictability des PI est 100 %, 83,33 %, 50 % et 66,67 %")
def _ratios(world):
    got = [r["FeaturePredictability"] for r in sheet(world, "PIMetrics")]
    assert all(close(value, expected) for value, expected in zip(got, (1.0, 5 / 6, 3 / 6, 4 / 6))), got


@then("la feuille IterationFeatures de l'ART est celle calculée à la main")
def _art_rows(world):
    rows = [r for r in sheet(world, "IterationFeatures") if r["Team"] == "ART"]
    assert [r["Iteration-label"] for r in rows] == list(ART_ROWS)
    got = {r["Iteration-label"]: (r["Engaged"], r["Added"], r["Deferred"], r["Released"], r["ReleasedAdded"], r["CarryOver"]) for r in rows}
    assert got == ART_ROWS
    assert {r["Iteration-label"]: r["PI"] for r in rows} == {label: "PI-" + label[3] for label in ART_ROWS}


@then("chaque équipe a ses lignes de PI-3 et PI-4 calculées à la main")
def _team_rows(world):
    rows = {(r["Team"], r["Iteration-label"]): (r["Engaged"], r["Added"], r["Deferred"], r["Released"]) for r in sheet(world, "IterationFeatures")}
    for team, expected in TEAM_ROWS.items():
        assert {label: rows[(team, label)] for label in expected} == expected, team


@then("la vélocité par itération est celle du scénario")
def _velocity(world):
    rows = {r["Iteration-label"]: r["Velocity"] for r in sheet(world, "Metrics") if r["Team"] == "ART"}
    assert rows == VELOCITY


@then("chaque itération qui livre a un Flow Time de 8 jours et le Flow Efficiency de son PI")
def _flow_time_and_efficiency(world):
    delivering = 0
    for r in (r for r in sheet(world, "Metrics") if r["Team"] == "ART"):
        label = r["Iteration-label"]
        if VELOCITY[label]:
            delivering += 1
            assert close(r["FlowTimeMedian"], FLOW_TIME_DAYS) and close(r["FlowTimeP85"], FLOW_TIME_DAYS), (label, r["FlowTimeMedian"])
            assert close(r["FlowEfficiency"], FLOW_EFFICIENCY["PI-" + label[3]]), (label, r["FlowEfficiency"])
            assert close(r["DistDoneFeature"], 1.0), label
        else:
            assert r["FlowTimeMedian"] == "not computable" and r["FlowEfficiency"] == "not computable", label
    assert delivering == sum(1 for v in VELOCITY.values() if v)


@then("chaque statut dure en moyenne 1 jour, 2 jours pour le statut long du PI de la Feature")
def _status_durations(world):
    durations: dict[tuple[str, str], list[float]] = {}
    for r in sheet(world, "StatusIntervals"):
        if r["LeftUtc"] in (None, ""):
            continue                                                       # the visit still open: todo is behind, deferred has no end
        number = int(r["Key"].rsplit("-", 1)[1])
        durations.setdefault((PI_OF_FEATURE[number], r["Status"]), []).append(float(r["DurationDays"]))
    assert {status for _pi, status in durations} >= set(CHAIN) | {"todo"}
    for (pi, status), days in durations.items():
        if status == "todo":
            continue                                                       # the wait in the backlog before the chain: not a cycle time
        expected = 2.0 if status == LONG_STATUS[pi] else 1.0
        assert all(abs(d - expected) < 1e-6 for d in days), (pi, status, days)
    assert ITERATION_STARTS
