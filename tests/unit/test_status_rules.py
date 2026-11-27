"""The status rules that come from Jira (flow/status_rules.py): the category, the start, the end and the free flag of every status, with the
exceptions the person typed in the sheet Workflows winning over what Jira says (AL, 2026-11-18: the metrics use the _from_Jira sheets, the
_from_user sheet only corrects what Jira cannot decide). Every figure is typed by hand.
"""
from jira2dashboard.flow.model import status_key
from jira2dashboard.flow.status_rules import (Exception_, KindCounts, Rule, counts_line, effective_rules, first_in_progress, has_wait, jira_category,
                                              rule_counts, rule_warnings, unknown_statuses)

NEW, IND, DONE = "new", "indeterminate", "done"


def rows(catalog, exceptions=(), kinds=("Feature",), starts=None, free=(), current=None):
    got = effective_rules(kinds, catalog, current or {}, set(free), list(exceptions), starts)
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
    got = rows({"Nouveau": NEW, "Analyse": IND, "Fini": DONE, "Parque": NEW}, free={"Parque"})
    assert got[("Feature", "Nouveau")] == Rule("Feature", "Nouveau", "Backlog", 0, 0, 0, "Jira")
    assert got[("Feature", "Analyse")] == Rule("Feature", "Analyse", "Active", 1, 0, 0, "Jira")          # in progress: it starts the flow
    assert got[("Feature", "Fini")] == Rule("Feature", "Fini", "Done", 0, 1, 0, "Jira")                 # every Done status ends the flow
    assert got[("Feature", "Parque")] == Rule("Feature", "Parque", "Backlog", 0, 0, 1, "Jira")


def test_an_exception_wins_field_by_field_and_an_empty_cell_takes_jiras_value():
    got = rows({"Analyse": IND, "Attente": IND, "Fini": DONE, "Annule": DONE, "Deploy": DONE},
               exceptions=[Exception_(None, "Attente", "Wait", None, None),
                           Exception_(None, "Annule", "Removed", None, None),
                           Exception_(None, "Deploy", "Wait", None, None),
                           Exception_(None, "Analyse", None, None, 1)])
    assert got[("Feature", "Attente")].category == "Wait" and got[("Feature", "Attente")].source == "user"
    assert (got[("Feature", "Annule")].category, got[("Feature", "Annule")].is_end) == ("Removed", 0)           # no longer Done: not an end
    assert (got[("Feature", "Deploy")].category, got[("Feature", "Deploy")].is_end) == ("Wait", 0)
    assert (got[("Feature", "Fini")].category, got[("Feature", "Fini")].is_end, got[("Feature", "Fini")].source) == ("Done", 1, "Jira")
    analyse = got[("Feature", "Analyse")]
    assert (analyse.category, analyse.is_start, analyse.is_free, analyse.source) == ("Active", 1, 1, "user")      # only IsFree was typed


def test_an_explicit_zero_or_one_beats_the_default_even_against_the_category():
    got = rows({"Fini": DONE, "Suite": IND}, exceptions=[Exception_(None, "Fini", None, 0, None), Exception_(None, "Suite", None, 1, None)])
    assert got[("Feature", "Fini")].is_end == 0 and got[("Feature", "Suite")].is_end == 1


def test_an_exception_of_one_workflow_only_touches_that_workflow_and_the_most_precise_row_wins():
    got = rows({"Test": IND}, kinds=("Feature", "Team"), exceptions=[Exception_(None, "Test", "Wait", None, None),
                                                                    Exception_("Team", "Test", "Active", None, None)])
    assert got[("Feature", "Test")].category == "Wait" and got[("Team", "Test")].category == "Active"


def test_the_most_precise_row_wins_whatever_the_order_of_the_rows():
    got = rows({"Test": IND}, kinds=("Feature", "Team"), exceptions=[Exception_("Team", "Test", "Active", None, None),
                                                                    Exception_(None, "Test", "Wait", None, None)])
    assert got[("Feature", "Test")].category == "Wait" and got[("Team", "Test")].category == "Active"


def test_a_status_that_only_the_exceptions_or_the_current_statuses_know_is_still_ruled():
    got = rows({"Fini": DONE}, current={"Vieux": IND}, exceptions=[Exception_(None, "Autre", "Wait", None, None)])
    assert got[("Feature", "Vieux")].category == "Active"                              # from the category Jira gives a ticket's current status
    assert got[("Feature", "Autre")].category == "Wait"
    assert rows({}, current={"Vieux": "whatever"})[("Feature", "Vieux")].category == "Unclassified"      # no source knows it: it stays visible


def test_the_catalog_beats_the_current_status_category_of_a_ticket():
    assert rows({"Etat": DONE}, current={"Etat": IND})[("Feature", "Etat")].category == "Done"


def test_the_rules_come_out_in_the_order_of_the_kinds_then_the_names():
    got = effective_rules(("Feature", "Team"), {"B": NEW, "A": NEW}, {}, set(), [])
    assert [(r.workflow, r.status) for r in got] == [("Feature", "A"), ("Feature", "B"), ("Team", "A"), ("Team", "B")]


def test_flow_efficiency_needs_a_declared_wait_status():
    assert has_wait(list(effective_rules(("Feature",), {"A": IND}, {}, set(), []))) is False
    assert has_wait(list(effective_rules(("Feature",), {"A": IND}, {}, set(), [Exception_(None, "A", "Wait", None, None)]))) is True


# --- one status whatever its spelling (AL, 2026-11-27: two spellings in the catalog of the site gave IsStart = 2 in Excel) -------------------

def test_the_key_of_a_status_is_its_name_trimmed_case_ignored_as_excel_compares():
    assert status_key(" En Analyse ") == status_key("en analyse") == "en analyse"
    assert status_key("Prêt") == status_key("PRÊT") and status_key("A déployer") != status_key("À déployer")       # an accent is no case
    assert status_key(None) == ""


def test_two_spellings_of_one_status_are_one_rule_named_as_the_catalog_spells_it_first():
    got = effective_rules(("Feature",), {"En analyse": NEW, "En Analyse": IND, "Fini": DONE}, {}, set(), [])
    assert [(r.status, r.category, r.is_start) for r in got] == [("En Analyse", "Active", 1), ("Fini", "Done", 0)]      # binary order: "A" < "a"


def test_an_exception_typed_in_another_case_or_with_spaces_rules_the_status_of_jira():
    got = effective_rules(("Feature",), {"Prêt": IND}, {}, set(), [Exception_("Feature", " prêt ", "Wait", None, None)])
    assert [(r.status, r.category, r.source) for r in got] == [("Prêt", "Wait", "user")]


def test_the_current_status_of_a_ticket_in_another_case_is_the_status_of_the_catalog():
    got = effective_rules(("Feature",), {"Terminé": DONE}, {"TERMINÉ": IND}, set(), [])
    assert [(r.status, r.category) for r in got] == [("Terminé", "Done")]


def test_the_rule_of_a_status_is_found_whatever_the_spelling_of_the_ticket():
    from jira2dashboard.flow.model import Rules
    from jira2dashboard.params.model import ParamSheet
    header = ("Workflow", "Status", "Category", "IsEnd", "IsFree")
    rules = Rules.from_params([ParamSheet("Workflows", header, ())], {"En Analyse": IND, "En analyse": IND}, {})
    assert rules.rule("Feature", "EN ANALYSE ").is_start and rules.rule("Feature", "en analyse").category == "Active"
    assert [r.status for r in rules.listed() if r.workflow == "Feature"] == ["En Analyse"]


# --- the start of the flow: a status in progress, whatever the workflow (AL, 2026-11-27) ----------------------------------------------

def test_every_status_in_progress_starts_the_flow_active_or_wait_and_nothing_else():
    got = rows({"Nouveau": NEW, "Analyse": IND, "Pret": IND, "Dev": IND, "Fini": DONE, "Annule": DONE},
               exceptions=[Exception_(None, "Pret", "Wait", None, None), Exception_(None, "Annule", "Removed", None, None)])
    assert {s for (_k, s), r in got.items() if r.is_start} == {"Analyse", "Pret", "Dev"}


def test_an_exception_backlog_on_a_status_in_progress_moves_the_start_later():
    got = rows({"Nouveau": NEW, "Analyse": IND, "Dev": IND}, exceptions=[Exception_("Feature", "Analyse", "Backlog", None, None)])
    assert (got[("Feature", "Analyse")].is_start, got[("Feature", "Dev")].is_start) == (0, 1)


def test_the_start_statuses_of_config_alone_start_the_flow_of_their_kind():
    got = rows({"Nouveau": NEW, "Analyse": IND, "Dev": IND, "Fini": DONE}, kinds=("Feature", "Team"), starts={"Feature": [" dev "], "Team": []})
    assert {s for (k, s), r in got.items() if r.is_start and k == "Feature"} == {"Dev"}                   # compared by key
    assert {s for (k, s), r in got.items() if r.is_start and k == "Team"} == {"Analyse", "Dev"}           # empty: the statuses in progress


def test_a_start_status_of_config_that_no_source_knows_starts_nothing():
    got = rows({"Analyse": IND}, starts={"Feature": ["Inconnu"]})
    assert not any(r.is_start for r in got.values())


def test_the_statuses_that_a_setting_names_and_jira_does_not_know_are_listed_once():
    exceptions = [Exception_("Feature", "À déployer", "Wait", None, None), Exception_("Team", "à déployer", "Wait", None, None),
                  Exception_(None, "prêt", "Wait", None, None)]
    got = unknown_statuses(exceptions, {"Feature": ["En dev", "Prêt"], "Team": []}, ["Prêt", "A déployer"])
    assert got == [("À déployer", "Workflows"), ("En dev", "FeatureStartStatuses")]
    assert unknown_statuses([], {}, []) == []


# --- what the log says about the rules, and the warnings when a metric cannot be computed (AL, 2026-11-22) ---------------------------

def test_the_counts_of_each_kind_name_its_starts_actives_waits_and_ends():
    got = effective_rules(("Feature", "Team"), {"Funnel": NEW, "Analyse": IND, "Dev": IND, "Done": DONE, "Cancelled": DONE}, {}, set(),
                          [Exception_("Team", "Dev", "Wait", None, None)])
    counts = rule_counts(got, ("Feature", "Team"))
    assert counts == [KindCounts("Feature", 2, 2, 0, 2), KindCounts("Team", 2, 1, 1, 2)]
    assert counts_line(counts) == "status rules | Feature: start 2, active 2, wait 0, end 2 | Team: start 2, active 1, wait 1, end 2"


def test_a_kind_with_tickets_and_no_start_status_is_warned_with_its_code():
    counts = [KindCounts("Feature", 1, 2, 1, 1), KindCounts("Team", 0, 3, 0, 1)]
    assert rule_warnings(counts, {"Feature", "Team"}) == [
        "J2D-FLOW-010 | no start status in the Team workflow: Flow Time, Flow Efficiency and Flow Load cannot be computed"
        " | the statuses in progress start the flow: check the Category of its statuses in the sheet Workflows_from_user and its start statuses in "
        "the sheet Config_from_user"]


def test_a_kind_without_tickets_is_not_warned_and_no_ticket_at_all_gives_no_warning():
    counts = [KindCounts("Feature", 0, 0, 0, 0), KindCounts("Team", 1, 1, 1, 1)]
    assert rule_warnings(counts, {"Team"}) == []
    assert rule_warnings([KindCounts("Feature", 0, 0, 0, 0), KindCounts("Team", 0, 0, 0, 0)], set()) == []


def test_no_wait_status_in_any_kind_is_warned_once_after_the_starts():
    counts = [KindCounts("Feature", 0, 1, 0, 1), KindCounts("Team", 1, 1, 0, 1)]
    assert rule_warnings(counts, {"Feature", "Team"}) == [
        "J2D-FLOW-010 | no start status in the Feature workflow: Flow Time, Flow Efficiency and Flow Load cannot be computed"
        " | the statuses in progress start the flow: check the Category of its statuses in the sheet Workflows_from_user and its start statuses in "
        "the sheet Config_from_user",
        "J2D-FLOW-011 | no status is declared Wait: Flow Efficiency cannot be computed"
        " | declare Category = Wait in the sheet Workflows_from_user for the statuses where the work waits"]


def test_the_vba_says_the_same_words():
    from pathlib import Path
    from jira2dashboard.flow import status_rules as R
    vba = (Path(__file__).resolve().parents[2] / "vba" / "modParams.bas").read_text(encoding="ascii")
    assert f'Private Const NO_START_TEXT As String = "{R.NO_START_TEXT}"' in vba
    assert f'Private Const NO_WAIT_TEXT As String = "{R.NO_WAIT_TEXT}"' in vba
    assert f'"{R.NO_START_CODE} | no start status in the " & CStr(kind) & NO_START_TEXT' in vba and f'"{R.NO_WAIT_CODE} | " & NO_WAIT_TEXT' in vba
    assert '"status rules" & out' in vba and '": start " & KindCount' in vba


def test_the_rules_of_the_python_path_are_counted_the_same_way():
    from jira2dashboard.flow.model import Rules, StatusRule
    rules = Rules({("Team", "Dev"): StatusRule("Active", True, False, False), ("Team", "Done"): StatusRule("Done", False, True, False),
                   ("Feature", "Wait"): StatusRule("Wait", False, False, False)})
    assert rule_counts(rules.listed(), ("Feature", "Team")) == [KindCounts("Feature", 0, 0, 1, 0), KindCounts("Team", 1, 1, 0, 1)]



def test_the_rules_keep_their_source_for_the_sheet_status_rules_feature_first():
    from jira2dashboard.flow.model import Rules
    from jira2dashboard.flow.status_rules import status_rule_rows
    rows = status_rule_rows(Rules.from_params(catalog={"Doing": "indeterminate"}).entries)
    assert rows and rows[0][0] == "Feature" and {r[6] for r in rows} <= {"Jira", "user"} and {r[7] for r in rows} == {""}
    kinds = [r[0] for r in rows]
    assert kinds == sorted(kinds, key=lambda k: k != "Feature")
