"""Steps of features/macro_check.feature: the macro check of `scripts\\run.ps1 check-machine`. Drives a real Excel: marker `excel`.

Only opens a workbook and runs a read-only function: no extraction, so it stays clear of the open crash defect.
"""
import pytest
from pytest_bdd import given, parsers, scenarios, then, when

from jira2dashboard import diagnose as D
from jira2dashboard.build_workbook import build_delivery
from jira2dashboard.excel import com
from jira2dashboard.paths import INTER_DIR

scenarios("../features/macro_check.feature")

SOURCE = INTER_DIR / "macro_check_source.xlsm"          # the workbook to deliver, built by the code of `run.ps1 build`
PLAIN = INTER_DIR / "macro_check_plain.xlsx"            # a workbook without any macro
MACRO = "modActions.ProgressSnapshot"


@pytest.fixture
def world():
    return {}


@given("le classeur à livrer")
def _delivered(world):
    build_delivery(SOURCE)
    world["path"] = SOURCE


@given("un classeur sans macro")
def _plain(world):
    with com.excel_session() as app:
        wb = com.new_workbook(app)
        wb.SaveAs(str(PLAIN.resolve()), 51)             # xlOpenXMLWorkbook
        wb.Close(False)
        del wb, app
    world["path"] = PLAIN


@when("je fais exécuter sa macro en lecture seule")
def _run(world):
    with com.excel_session() as app:
        world["ran"], world["detail"] = com.try_workbook_macro(app, world["path"], MACRO)
        del app


@then("la macro s'exécute et rend l'état du tableau de bord")
def _ran(world):
    assert world["ran"], world["detail"]
    assert "|" in world["detail"]                       # ProgressSnapshot gives "fraction|status text"


@then("la macro ne s'exécute pas")
def _not_ran(world):
    assert world["ran"] is False and world["detail"]


@then("le conseil parle de la protection du poste")
def _advice(world):
    readings = D.Readings(office={}, policies={}, trusted_locations=[], security_products=None, interpreters=[])
    assert "endpoint protection" in " ".join(D.macro_advice(world["ran"], world["path"], readings))
