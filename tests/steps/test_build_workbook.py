"""Steps of features/build_workbook.feature (walking skeleton). Drives a real Excel: marker `excel`."""
import pytest
from pytest_bdd import given, scenarios, then, when

from jira2dashboard.paths import INTER_DIR
from jira2dashboard.testbench import run_vba_tests as rvt

scenarios("../features/build_workbook.feature")

TARGET = INTER_DIR / "walking_skeleton.xlsm"


@pytest.fixture
def world():
    """State shared by the steps of one scenario."""
    return {}


@given("les modules VBA du repo")
def _modules(world):
    world["names"] = rvt.ORDER


@when("je construis le classeur de test")
def _build(world):
    world["markers"] = rvt.build(world["names"], TARGET)


@when("je lance les tests VBA avec le réseau")
def _run_with_network(world):
    world["report"] = rvt.run_tests(TARGET, network=True)


@then("le classeur compile")
def _compiles(world):
    assert rvt.compile_ok(TARGET) == "COMPILED_OK"


@then("les tests VBA hors ligne passent sans échec")
def _offline_tests_pass(world):
    report = rvt.run_tests(TARGET, network=False)
    assert report["ko"] == []
    assert report["total"] > 0


@then("tous les marqueurs de couverture sont atteints")
def _all_markers_hit(world):
    hits = set(world["report"]["hits"])
    missed = [m for m in world["markers"] if m not in hits]
    assert world["report"]["ko"] == []
    assert not missed, missed
