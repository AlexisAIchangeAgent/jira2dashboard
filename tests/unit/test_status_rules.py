"""The status rules that come from Jira (flow/status_rules.py): the category, the start, the end and the free flag of every status, with the
exceptions the person typed in the sheet Workflows winning over what Jira says (AL, 2026-11-18: the metrics use the _from_Jira sheets, the
_from_user sheet only corrects what Jira cannot decide). Every figure is typed by hand.
"""
from jira2dashboard.flow.status_rules import Exception_, Rule, effective_rules, first_in_progress, has_wait, jira_category

NEW, IND, DONE = "new", "indeterminate", "done"


def rows(catalog, exceptions=(), kinds=("Feature",), starts=(), free=(), current=None):
    got = effective_rules(kinds, catalog, current or {}, set(starts), set(free), list(exceptions))
    return {(r.workflow, r.status): r for r in got}


# --- the Jira categories --------------------------------------------------------------------------------------------------------

def test_the_three_categories_of_jira_in_all_the_spellings_the_apis_use():
    assert [jira_category(k) for k in ("new", "NEW", "TODO", "To Do", "indeterminate", "IN_PROGRESS", "In Progress", "done", "DONE", "Done")] == [
        "Backlog", "Backlog", "Backlog", "Backlog", "Active", "Active", "Active", "Done", "Done", "Done"]
    assert jira_category("") is None and jira_category(None) is None and jira_category("whatever") is None


# --- the first status in progress, from the status of creation ----------------------------------------------------------------------

CATEGORIES = {"Nouveau": NEW, "Analyse": IND, "Pret": IND, "Dev": IND, "Fini": DONE, "Idee": NEW}


def test_the_start_is_the_first_in_progress_status_reached_from_the_creation_status():
    moves = [("Nouveau", "Analyse"), ("Analyse", "Pret"), ("Pret", "Dev"), ("Dev", "Fini")]
    assert first_in_progress("Nouveau", moves, CATEGORIES) == "Analyse"


def test_a_status_in_progress_that_is_nearer_wins_and_the_first_listed_breaks_a_tie():
    moves = [("Nouveau", "Idee"), ("Idee", "Dev"), ("Nouveau", "Pret"), ("Nouveau", "Analyse")]
    assert first_in_progress("Nouveau", moves, CATEGORIES) == "Pret"              # depth 1: Pret and Analyse, Pret comes first; Dev is at depth 2


def test_loops_and_statuses_that_cannot_be_reached_do_not_stop_the_search():
    moves = [("Nouveau", "Idee"), ("Idee", "Nouveau"), ("Fini", "Dev")]            # Dev is only reachable from Fini, which nothing leads to
    assert first_in_progress("Nouveau", moves, CATEGORIES) is None
    assert first_in_progress(None, moves, CATEGORIES) is None and first_in_progress("Nouveau", [], CATEGORIES) is None


def test_a_creation_status_that_is_already_in_progress_is_the_start():
    assert first_in_progress("Dev", [("Dev", "Fini")], CATEGORIES) == "Dev"


# --- the effective rules ---------------------------------------------------------------------------------------------------------

def test_without_any_exception_jira_gives_the_category_the_end_the_start_and_the_free_flag():
    got = rows({"Nouveau": NEW, "Analyse": IND, "Fini": DONE, "Parque": NEW}, starts={"Analyse"}, free={"Parque"})
    assert got[("Feature", "Nouveau")] == Rule("Feature", "Nouveau", "Backlog", 0, 0, 0, "Jira")
    assert got[("Feature", "Analyse")] == Rule("Feature", "Analyse", "Active", 1, 0, 0, "Jira")
    assert got[("Feature", "Fini")] == Rule("Feature", "Fini", "Done", 0, 1, 0, "Jira")                 # every Done status ends the flow
    assert got[("Feature", "Parque")] == Rule("Feature", "Parque", "Backlog", 0, 0, 1, "Jira")


def test_an_exception_wins_field_by_field_and_an_empty_cell_takes_jiras_value():
    got = rows({"Analyse": IND, "Attente": IND, "Fini": DONE, "Annule": DONE, "Deploy": DONE}, starts={"Analyse"},
               exceptions=[Exception_(None, "Attente", "Wait", None, None, None),
                           Exception_(None, "Annule", "Removed", None, None, None),
                           Exception_(None, "Deploy", "Wait", None, None, None),
                           Exception_(None, "Analyse", None, None, None, 1)])
    assert got[("Feature", "Attente")].category == "Wait" and got[("Feature", "Attente")].source == "user"
    assert (got[("Feature", "Annule")].category, got[("Feature", "Annule")].is_end) == ("Removed", 0)           # no longer Done: not an end
    assert (got[("Feature", "Deploy")].category, got[("Feature", "Deploy")].is_end) == ("Wait", 0)
    assert (got[("Feature", "Fini")].category, got[("Feature", "Fini")].is_end, got[("Feature", "Fini")].source) == ("Done", 1, "Jira")
    analyse = got[("Feature", "Analyse")]
    assert (analyse.category, analyse.is_start, analyse.is_free, analyse.source) == ("Active", 1, 1, "user")      # only IsFree was typed


def test_an_explicit_zero_or_one_beats_the_default_even_against_the_category():
    got = rows({"Fini": DONE, "Suite": IND}, exceptions=[Exception_(None, "Fini", None, None, 0, None), Exception_(None, "Suite", None, None, 1, None)])
    assert got[("Feature", "Fini")].is_end == 0 and got[("Feature", "Suite")].is_end == 1


def test_an_exception_of_one_workflow_only_touches_that_workflow_and_the_most_precise_row_wins():
    got = rows({"Test": IND}, kinds=("Feature", "Team"), exceptions=[Exception_(None, "Test", "Wait", None, None, None),
                                                                    Exception_("Team", "Test", "Active", None, None, None)])
    assert got[("Feature", "Test")].category == "Wait" and got[("Team", "Test")].category == "Active"


def test_the_most_precise_row_wins_whatever_the_order_of_the_rows():
    got = rows({"Test": IND}, kinds=("Feature", "Team"), exceptions=[Exception_("Team", "Test", "Active", None, None, None),
                                                                    Exception_(None, "Test", "Wait", None, None, None)])
    assert got[("Feature", "Test")].category == "Wait" and got[("Team", "Test")].category == "Active"


def test_a_status_that_only_the_exceptions_or_the_current_statuses_know_is_still_ruled():
    got = rows({"Fini": DONE}, current={"Vieux": IND}, exceptions=[Exception_(None, "Autre", "Wait", None, None, None)])
    assert got[("Feature", "Vieux")].category == "Active"                              # from the category Jira gives a ticket's current status
    assert got[("Feature", "Autre")].category == "Wait"
    assert rows({}, current={"Vieux": "whatever"})[("Feature", "Vieux")].category == "Unclassified"      # no source knows it: it stays visible


def test_the_catalog_beats_the_current_status_category_of_a_ticket():
    assert rows({"Etat": DONE}, current={"Etat": IND})[("Feature", "Etat")].category == "Done"


def test_the_rules_come_out_in_the_order_of_the_kinds_then_the_names():
    got = effective_rules(("Feature", "Team"), {"B": NEW, "A": NEW}, {}, set(), set(), [])
    assert [(r.workflow, r.status) for r in got] == [("Feature", "A"), ("Feature", "B"), ("Team", "A"), ("Team", "B")]


def test_flow_efficiency_needs_a_declared_wait_status():
    assert has_wait(list(effective_rules(("Feature",), {"A": IND}, {}, set(), set(), []))) is False
    assert has_wait(list(effective_rules(("Feature",), {"A": IND}, {}, set(), set(), [Exception_(None, "A", "Wait", None, None, None)]))) is True
