"""Steps of features/flow_metrics.feature: the reference implementation on the hand-written cases (no Excel)."""
import pytest
from pytest_bdd import given, parsers, scenarios, then, when

from conftest import FLOW_AS_OF, at
from jira2dashboard.flow.itemflow import item_flow
from jira2dashboard.flow.load import daily_load
from jira2dashboard.flow.metrics import team_period_metrics
from jira2dashboard.flow.model import NOT_COMPUTABLE, Period, Rules

scenarios("../features/flow_metrics.feature")

PERIODS = {"P1": Period("P1", at("10-05"), at("10-19")), "P2": Period("P2", at("10-19"), at("11-02"))}
METRICS = {
    "Flow Velocity": "velocity", "Flow Time médiane": "flow_time_median", "Flow Time P85": "flow_time_p85",
    "Flow Time exclus": "flow_time_excluded", "Flow Efficiency": "flow_efficiency",
    "Flow Load moyenne": "flow_load_average", "Flow Load fin de période": "flow_load_end",
    "Flow Predictability": "flow_predictability",
}


@pytest.fixture
def world():
    return {}


@given("les neuf tickets de flux écrits à la main")
def _cases(world, flow_cases):
    rules = Rules.from_params()
    world["items"] = [item_flow(i, FLOW_AS_OF, rules) for i in flow_cases]


@when(parsers.parse('je calcule la ligne du ticket "{key}"'))
def _row(world, key):
    world["row"] = next(i for i in world["items"] if i.key == key)


@when(parsers.parse('je calcule les métriques de l\'équipe "{team}" pour la période "{period}" en mode "{mode}"'))
def _metrics(world, team, period, mode):
    world["metrics"] = team_period_metrics(world["items"], team, PERIODS[period], mode, FLOW_AS_OF)


@when(parsers.parse('je calcule le Flow Load de l\'équipe "{team}" en mode "{mode}" pour la période "{period}"'))
def _load(world, team, mode, period):
    p = PERIODS[period]
    world["load"] = daily_load(world["items"], mode, team, p.start, p.end, FLOW_AS_OF)


@then(parsers.parse('son démarrage est le "{start}" et son premier démarrage le "{first}"'))
def _start(world, start, first):
    assert (world["row"].start, world["row"].first_start) == (at(start), at(first))


@then(parsers.parse("son Flow Time vaut {days:d} jours"))
def _flow_time(world, days):
    assert world["row"].flow_time_days == days


@then("son Flow Time est vide")
def _flow_time_empty(world):
    assert world["row"].flow_time_days is None


@then(parsers.parse("il a {count:d} retours en backlog"))
def _backlog_returns(world, count):
    assert world["row"].backlog_returns == count


@then("il est un démarrage sauté")
def _skipped(world):
    assert world["row"].skipped_start and world["row"].excluded_from_flow_time


@then(parsers.parse("il a été rouvert {count:d} fois"))
def _reopened(world, count):
    assert world["row"].reopened == count and world["row"].excluded_from_flow_time


@then(parsers.parse('la métrique "{name}" vaut "{value}"'))
def _metric(world, name, value):
    actual = getattr(world["metrics"], METRICS[name])
    if value == NOT_COMPUTABLE:
        assert actual == NOT_COMPUTABLE
    else:
        assert actual != NOT_COMPUTABLE and actual == pytest.approx(float(value), abs=1e-9 if "." not in value else 1e-6)


@then(parsers.parse('la charge du "{day}" vaut {load:d}'))
def _daily_load(world, day, load):
    assert world["load"][at(day)] == load


@then(parsers.parse('la part "{work_type}" des tickets terminés vaut {percent:d} %'))
def _share(world, work_type, percent):
    assert world["metrics"].distribution_done[work_type] == pytest.approx(percent / 100)


@then(parsers.parse("la somme des parts des tickets terminés vaut {percent:d} %"))
def _shares_sum(world, percent):
    assert sum(world["metrics"].distribution_done.values()) == pytest.approx(percent / 100)


@then("les tickets retenus plus les tickets exclus font la vitesse de flux")
def _included_plus_excluded(world):
    m = world["metrics"]
    assert m.flow_time_count + m.flow_time_excluded == m.velocity > 0
