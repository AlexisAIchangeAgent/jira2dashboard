"""Steps of features/params.feature. Fast scenarios are pure; `@excel` scenarios share one built workbook."""
import shutil

import pytest
from pytest_bdd import given, parsers, scenarios, then, when

from jira2dashboard import sheet_names as S
from jira2dashboard.params import model
from jira2dashboard.params.write import EXAMPLES
from jira2dashboard.paths import PARAMS_DIR

scenarios("../features/params.feature")

VERDICTS = {"autorisée": True, "refusée": False}


@pytest.fixture
def world():
    """State shared by the steps of one scenario."""
    return {}


def _normalize(value):
    """Excel returns whole numbers as floats."""
    return int(value) if isinstance(value, float) and value.is_integer() else value


# --- pure scenarios ---------------------------------------------------------------------------------

@given("les fichiers de paramètres du dépôt")
def _params_copy(world, tmp_path):
    world["dir"] = tmp_path / "params"
    shutil.copytree(PARAMS_DIR, world["dir"])


@when(parsers.parse('je remplace "{old}" par "{new}" dans "{filename}"'))
def _replace(world, old, new, filename):
    path = world["dir"] / filename
    text = path.read_text(encoding="utf-8")
    assert old in text
    path.write_text(text.replace(old, new, 1), encoding="utf-8")


@when("je les charge")
def _load(world):
    try:
        world["sheets"] = {s.name: s for s in model.load_all(world["dir"])}
        world["error"] = None
    except model.ParamsError as err:
        world["error"] = str(err)


@then(parsers.parse('la feuille "{name}" a {count:d} lignes'))
def _row_count(world, name, count):
    assert world["error"] is None, world["error"]
    assert len(world["sheets"][name].rows) == count


@then(parsers.parse('les paramètres obligatoires de la feuille Config sont "{names}"'))
def _mandatory(world, names):
    assert world["error"] is None, world["error"]
    assert ", ".join(model.mandatory_parameters(world["sheets"]["Config"])) == names


@then(parsers.parse('les colonnes de la feuille Config sont "{names}"'))
def _config_columns(world, names):
    assert world["error"] is None, world["error"]
    assert ", ".join(world["sheets"]["Config"].header) == names


@then("chaque paramètre de la feuille Config a un exemple")
def _examples(world):
    config = world["sheets"]["Config"]
    column = config.header.index("Example")
    assert [row[0] for row in config.rows if not row[column]] == []


@then(parsers.parse('le chargement est refusé avec "{text}"'))
def _refused(world, text):
    assert world["error"] is not None and text in world["error"], world["error"]


# --- Excel scenarios --------------------------------------------------------------------------------

@given("le classeur construit avec les paramètres du dépôt")
def _workbook(world, built_workbook):
    world["wb"] = built_workbook["wb"]


@then(parsers.parse('la feuille Excel "{sheet}" reproduit le fichier "{filename}"'))
def _sheet_matches_file(world, sheet, filename):
    expected = next(s for s in model.load_all() if s.name == sheet)
    assert (PARAMS_DIR / filename).is_file()
    values = world["wb"].Worksheets(S.of(sheet)).UsedRange.Value
    width = len(expected.header)
    actual = [tuple(_normalize(v) for v in row[:width]) for row in values]
    if sheet in EXAMPLES:      # Iterations and PIs: the column Example sits next to the table, in rows the file leaves empty
        assert [_normalize(row[width]) for row in values] == ["Example", *EXAMPLES[sheet]]
        actual = [row for row in actual if any(v is not None for v in row)]
    assert actual[0] == expected.header
    assert actual[1:] == list(expected.rows)


def _run(world, macro, *args):
    wb = world["wb"]
    module = "modTests" if macro in ("CategoryOf", "IsAllowed") else "modParams"     # these two moved: no caller in the delivered workbook
    return wb.Application.Run(f"'{wb.Name}'!{module}.{macro}", *args)


@when(parsers.parse('le VBA donne la catégorie du statut "{status}" du workflow "{workflow}"'))
def _vba_category(world, status, workflow):
    world["actual"] = _run(world, "CategoryOf", workflow, status)


@then(parsers.parse('la catégorie est "{category}"'))
def _category_is(world, category):
    assert world["actual"] == category


@when(parsers.parse('le VBA teste la transition de "{source}" vers "{target}" dans le workflow "{workflow}"'))
def _vba_transition(world, source, target, workflow):
    world["actual"] = _run(world, "IsAllowed", workflow, source, target)


@then(parsers.parse("la transition est {verdict}"))
def _transition_is(world, verdict):
    assert world["actual"] is VERDICTS[verdict]
