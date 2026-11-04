"""params\\*.md must say the same thing as docs/safe-flow-metrics-and-data-checks.md (the documented decisions)."""
import pytest

from jira2dashboard.params import model
from jira2dashboard.params.mdtable import parse_tables
from jira2dashboard.paths import REPO_ROOT

DOC = REPO_ROOT / "docs" / "safe-flow-metrics-and-data-checks.md"


@pytest.fixture(scope="module")
def doc_tables():
    return parse_tables(DOC.read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def sheets():
    return {s.name: s for s in model.load_all()}


def doc_table(doc_tables, first_header_cells):
    return next(t for t in doc_tables if t.header[:len(first_header_cells)] == first_header_cells)


def test_workflows_match_the_status_table_of_the_doc(doc_tables, sheets):
    table = doc_table(doc_tables, ("Order", "Team status", "Feature status", "Category"))
    expected = {"Team": [], "Feature": []}
    for order, team, feature, category in (r[:4] for r in table.rows):
        number = None if order == "-" else int(order)
        expected["Team"].append((number, team, category.split()[0]))
        expected["Feature"].append((number, feature, category.split()[0]))
    for workflow in ("Team", "Feature"):
        actual = [(r[1], r[2], r[3]) for r in sheets["Workflows"].rows if r[0] == workflow]
        assert actual == expected[workflow], workflow


def test_flags_match_the_documented_decisions(sheets):
    """IsStart on En analyse, IsEnd on Terminé only, IsFree on Dépriorisé (doc, section 1)."""
    for _wf, _order, status, _cat, start, end, free in sheets["Workflows"].rows:
        assert start == (status == "En analyse"), status
        assert end == (status == "Terminé"), status
        assert free == (status == "Dépriorisé"), status


def test_transitions_match_the_transition_table_of_the_doc(doc_tables, sheets):
    table = doc_table(doc_tables, ("From", "To", "Type"))
    assert len(table.rows) == 13
    for workflow, rename in (("Team", lambda s: s), ("Feature", lambda s: "En test / recette" if s == "En test" else s)):
        expected = [(rename(f), rename(t), kind) for f, t, kind in table.rows]
        actual = [r[1:] for r in sheets["Transitions"].rows if r[0] == workflow]
        assert actual == expected, workflow


def test_field_mapping_matches_the_table_of_the_doc(doc_tables, sheets):
    table = doc_table(doc_tables, ("Extract", "Logical", "Type", "CanBeEmpty"))
    assert tuple(sheets["FieldMapping"].rows) == tuple(table.rows)
    assert sheets["FieldMapping"].header == table.header[:7] + ("TypeConfirmed",)
