"""Steps of features/diagnostic.feature: the report of `scripts\\run.ps1 diagnose` (pure part: nothing is read from Windows)."""
import hashlib
from pathlib import Path

import pytest
from pytest_bdd import given, parsers, scenarios, then, when

from jira2dashboard import diagnose as D

scenarios("../features/diagnostic.feature")

LOG_LINES = ["2026-11-06 10:00:00.000 RUN button Extract from Jira and compute | Excel 16.0 build 1 64-bit | test.xlsm",
             "2026-11-06 10:00:01.000 BEGIN search page 1"]


@pytest.fixture
def world(tmp_path):
    home = tmp_path / "Users" / "someone"
    local = home / "AppData" / "Local"
    (local / "jira2dashboard" / "logs").mkdir(parents=True)
    return {"home": home, "local": local, "workbook": None, "report": ""}


def new_workbook(world, content=b"workbook bytes"):
    path = world["home"] / "Downloads" / "jira2dashboard.xlsm"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(content)
    world["workbook"], world["content"] = path, content
    return path


@given("un classeur téléchargé d'Internet (marque du web, zone 3)")
def _downloaded(world):
    path = new_workbook(world)
    Path(f"{path}:Zone.Identifier").write_text("[ZoneTransfer]\r\nZoneId=3\r\n", encoding="ascii")


@given(parsers.parse('un journal VBA et un python.log qui contient le jeton "{token}"'))
def _logs(world, token):
    folder = world["local"] / "jira2dashboard" / "logs"
    (folder / "jira2dashboard.log").write_text("\n".join(LOG_LINES) + "\n", encoding="ascii")
    (folder / "python.log").write_text(f"2026-11-06 10:00:02 Authorization: Basic {token}\n", encoding="ascii")
    new_workbook(world)


@given("un classeur dans le dossier de l'utilisateur")
def _in_home(world):
    new_workbook(world)


@given("aucun classeur")
def _none(world):
    world["workbook"] = None


@when("je lance le diagnostic")
def _run(world):
    readings = D.Readings(office={"Version": "16.0.1", "Platform": "x64"}, policies={"VBAWarnings (user)": "not set"},
                          trusted_locations=[], security_products=["Some Security Product"], interpreters=[])
    world["report"] = D.report(world["workbook"], readings, home=str(world["home"]), local_app_data=str(world["local"]),
                               temp=str(world["local"] / "Temp"))


@then(parsers.parse('le rapport indique la marque du web "{answer}" et la zone {zone:d}'))
def _motw(world, answer, zone):
    assert f"Mark of the Web: {answer} (ZoneId={zone}" in world["report"]


@then("le rapport donne le MD5 et le SHA-256 du classeur")
def _hashes(world):
    assert f"MD5: {hashlib.md5(world['content']).hexdigest()}" in world["report"]
    assert f"SHA-256: {hashlib.sha256(world['content']).hexdigest()}" in world["report"]


@then("le rapport cite la fin du journal VBA")
def _log_quoted(world):
    assert LOG_LINES[-1] in world["report"]


@then(parsers.parse('le rapport ne contient pas "{text}"'))
def _absent(world, text):
    assert text not in world["report"]


@then(parsers.parse('le rapport contient "{text}"'))
def _present(world, text):
    assert text in world["report"]


@then("le rapport ne contient pas le dossier de l'utilisateur")
def _no_home(world):
    assert str(world["home"]).lower() not in world["report"].lower()


@then("le rapport dit qu'il n'y a pas de classeur et comment en donner un")
def _no_workbook(world):
    assert "No workbook found" in world["report"] and "--workbook" in world["report"]
