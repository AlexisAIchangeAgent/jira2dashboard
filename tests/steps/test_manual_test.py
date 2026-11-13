"""Steps of features/manual_test.feature: the copy of scripts\\manual-test.ps1. Drives a real Excel: marker `excel`.

No extraction here (the copy is only read), so the scenario stays clear of the open crash defect of the extraction.
"""
import pytest
from pytest_bdd import given, parsers, scenarios, then, when

from jira2dashboard import sheet_names as S
from jira2dashboard import mock_workbook as M
from jira2dashboard.build_workbook import build_delivery
from jira2dashboard.excel import com
from jira2dashboard.paths import INTER_DIR

scenarios("../features/manual_test.feature")

SOURCE = INTER_DIR / "manual_source.xlsm"      # the workbook to deliver, built by the code of `run.ps1 build`
COPY = INTER_DIR / "manual_copy.xlsm"          # not the copy of manual-test.ps1, which AL may have open
CLIENT_VALUES = ("BaseUrl", "Email", "ApiToken", "AsOf")
TOLERANCE = 1e-6


@pytest.fixture
def world():
    return {}


def rows(ws):
    return [row for row in ws.Range(M.ROWS).Value2 if row[0] not in (None, "")]


def read(path):
    """{Parameter: (Value, Mandatory)} of Config, and the rows of Iterations and PIs, of the workbook at `path`."""
    with com.excel_session() as app:
        wb = com.open_without_macros(app, path)
        header, *table = wb.Worksheets(S.of("Config")).UsedRange.Value
        col = {name: n for n, name in enumerate(header)}
        config = {r[col["Parameter"]]: (r[col["Value"]], r[col["Mandatory"]]) for r in table if r[col["Parameter"]]}
        result = {"config": config, "Iterations": rows(wb.Worksheets(S.of("Iterations"))), "PIs": rows(wb.Worksheets(S.of("PIs")))}
        wb.Close(False)
        del wb, app
    return result


def same_rows(excel, iterations) -> bool:
    return len(excel) == len(iterations) and all(
        name == p.name and abs(start - com.to_excel_serial(p.start)) < TOLERANCE
        and abs(end - com.to_excel_serial(p.end)) < TOLERANCE for (name, start, end), p in zip(excel, iterations))


@given("le classeur à livrer")
def _delivered():
    build_delivery(SOURCE)


@when(parsers.parse("je prépare la copie d'essai manuel pour le port {port:d}"))
def _prepare(world, monkeypatch, port):
    monkeypatch.setattr(M, "COPY", COPY)
    assert M.main(["--workbook", str(SOURCE), "--port", str(port)]) == 0
    world["copy"], world["source"] = read(COPY), read(SOURCE)


@then(parsers.parse('la feuille Config de la copie pointe sur "{url}" avec le compte, la date de référence et les '
                    'champs du faux Jira'))
def _config(world, url):
    config = world["copy"]["config"]
    assert {name: config[name][0] for name in M.config_values(url)} == M.config_values(url)


@then("chaque paramètre obligatoire de la copie a une valeur")
def _mandatory(world):
    config = world["copy"]["config"]
    mandatory = [name for name, (_value, flag) in config.items() if str(flag).strip().lower() == "true"]
    assert set(mandatory) >= set(CLIENT_VALUES)
    assert [name for name in mandatory if config[name][0] in (None, "")] == []


@then("les feuilles Iterations et PIs de la copie ont les itérations et le PI du faux Jira")
def _iterations(world):
    assert same_rows(world["copy"]["Iterations"], M.iterations())
    assert same_rows(world["copy"]["PIs"], M.PIS)


@then("le classeur à livrer garde BaseUrl, Email, ApiToken et AsOf vides")
def _source_untouched(world):
    config = world["source"]["config"]
    assert {name: config[name][0] for name in CLIENT_VALUES} == dict.fromkeys(CLIENT_VALUES, None)
