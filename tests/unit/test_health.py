"""What the person should look at after a run (flow/health.py, the Python twin of modHealth.HealthFields): the same fields, the same words, each one
only when there is something to see. AL, 2026-11-22: "flow metrics are still not computable", so the fields say WHY (no start status, no Wait status,
AsOf before the data, a start in the second of the backlog entry...). Every figure is typed by hand from the small tickets below.
"""
import datetime as dt
import re
from pathlib import Path

from jira2dashboard import sheet_names as S
from jira2dashboard.flow import health as H
from jira2dashboard.flow.itemflow import item_flow
from jira2dashboard.flow.model import ART, OTHER_TYPES, Change, FlowIssue, Iteration, Rules, StatusRule
from jira2dashboard.flow.velocity_check import VelocityCheck

I1 = Iteration("I1", dt.datetime(2025, 5, 1), dt.datetime(2025, 5, 11))
AS_OF = dt.datetime(2025, 6, 1)
# Team: To Do (Backlog) -> Dev (Active, start) -> Review (Wait) -> Done (end). Feature: Funnel (Backlog) -> Build (Active, start) -> Shipped (end).
RULES = Rules({("Team", "To Do"): StatusRule("Backlog", False, False, False), ("Team", "Dev"): StatusRule("Active", True, False, False),
               ("Team", "Review"): StatusRule("Wait", False, False, False), ("Team", "Done"): StatusRule("Done", False, True, False),
               ("Feature", "Funnel"): StatusRule("Backlog", False, False, False), ("Feature", "Build"): StatusRule("Active", True, False, False),
               ("Feature", "Shipped"): StatusRule("Done", False, True, False)})


def day(text):
    return dt.datetime.fromisoformat(text)


def issue(key, moves, workflow="Team", team="Alpha", type_="Story", first=None, created="2025-04-01T00:00"):
    changes, previous = [], first or ("To Do" if workflow == "Team" else "Funnel")
    for n, (at, status) in enumerate(moves, start=1):
        changes.append(Change(f"{key}-{n}", day(at), previous, status))
        previous = status
    return FlowIssue(key=key, team=team, type=type_, work_type=None, workflow=workflow, created=day(created), status=previous, changes=tuple(changes))


def fields(issues, rules=RULES, as_of=AS_OF, checks=(), verdicts=(), feature_types="Feature", undated_pis=()):
    items = [item_flow(i, as_of, rules) for i in issues]
    return H.health_fields(issues, items, rules, [I1], as_of, list(checks), list(verdicts), feature_types, undated_pis)


CLEAN = [issue("T-1", [("2025-05-02T00:00", "Dev"), ("2025-05-03T00:00", "Review"), ("2025-05-04T00:00", "Done")]),
         issue("F-1", [("2025-05-02T00:00", "Build"), ("2025-05-05T00:00", "Shipped")], workflow="Feature", team=ART, type_="Feature")]


def test_a_clean_run_says_nothing():
    assert fields(CLEAN) == []


def test_a_kind_with_tickets_and_no_start_status_is_named():
    rules = Rules({k: StatusRule(r.category, False, r.is_end, False) if k[0] == "Team" else r for k, r in RULES._rows.items()})
    assert fields(CLEAN, rules) == [f" | No start status: Team (no Flow Time, Flow Efficiency or Flow Load): no status in progress, or a start status "
                                    f"that Jira does not know in {S.of('Config')}"]


def test_both_kinds_without_a_start_are_named_in_one_field_and_a_kind_without_tickets_is_not():
    rules = Rules({k: StatusRule(r.category, False, r.is_end, False) for k, r in RULES._rows.items()})
    assert fields(CLEAN, rules)[0].startswith(" | No start status: Feature, Team (")
    assert fields(CLEAN[:1], rules)[0].startswith(" | No start status: Team (")              # no Feature ticket: the Feature workflow does not matter


def test_no_wait_status_is_told():
    rules = Rules({k: StatusRule("Active" if r.category == "Wait" else r.category, r.is_start, r.is_end, False) for k, r in RULES._rows.items()})
    assert fields(CLEAN, rules) == [f" | No Wait status (no Flow Efficiency): set Category = Wait on the statuses where the work waits in "
                                    f"{S.of('Workflows')}"]


def test_the_tickets_of_the_program_that_are_not_features_and_a_feature_types_that_matches_none():
    sub = issue("F-2", [("2025-05-02T00:00", "Build")], workflow="Feature", team=OTHER_TYPES, type_="Sub-task")
    assert fields(CLEAN + [sub]) == [f" | Not Features: 1 of 2 tickets of the Program (team {OTHER_TYPES} in {S.of('ItemFlow')}, left out of the ART metrics)"]
    epic = issue("F-3", [], workflow="Feature", team=OTHER_TYPES, type_="Epic")
    got = fields(CLEAN[:1] + [sub, epic], feature_types="Fonctionnalite")
    assert got == [f" | Not Features: 2 of 2 tickets of the Program (team {OTHER_TYPES} in {S.of('ItemFlow')}, left out of the ART metrics)"     # one field
                   f' | FeatureTypes "Fonctionnalite" matches no ticket of the Program (types read: Sub-task, Epic): type the name of the Feature '
                   f"type of Jira in {S.of('Config')}"]


def test_the_tickets_done_outside_the_iterations_but_not_those_of_the_other_types():
    late = issue("T-2", [("2025-05-12T00:00", "Dev"), ("2025-05-13T00:00", "Done")])
    sub = issue("F-2", [("2025-05-12T00:00", "Build"), ("2025-05-13T00:00", "Shipped")], workflow="Feature", team=OTHER_TYPES, type_="Sub-task")
    got = fields(CLEAN + [late, sub])
    assert f" | Done outside the iterations: 1 of 3 done tickets (no metric counts them): add iterations to {S.of('Iterations')}" in got


def test_the_pis_without_dates_are_named_left_out_of_the_metrics():
    """AL, 2026-11-25: a PI of the sheet PIs with no Start and no End no longer stops 2. Compute the metrics (J2D-CFG-013): it is left out, and said."""
    names = ("PI#11-2027", "PI#3-2024", "PI#4.1-2024", "PI#4.2-2024", "PI#5-2025", "PI#6-2025")
    assert fields(CLEAN, undated_pis=names[:1]) == [
        f" | PIs without dates: 1 (PI#11-2027): left out of the metrics, type their Start and End in {S.of('PIs')}"]
    assert fields(CLEAN, undated_pis=names) == [f" | PIs without dates: 6 (PI#11-2027, PI#3-2024, PI#4.1-2024, PI#4.2-2024, PI#5-2025, ...): "
                                                f"left out of the metrics, type their Start and End in {S.of('PIs')}"]


def test_the_tickets_without_a_team():
    lost = issue("X-1", [], team="")
    assert fields(CLEAN + [lost]) == [f" | Without a team: 1 tickets (a ticket takes its team from {S.of('Projects')})"]


def test_the_statuses_no_rule_knows_are_named_once_in_their_order():
    odd = [issue(f"T-{n}", [("2025-05-02T00:00", s)]) for n, s in enumerate(("Odd", "Strange", "Odd", "A", "B", "C", "D"), start=10)]
    assert fields(CLEAN + odd) == [f" | Statuses not in Workflows: 6 (Odd, Strange, A, B, C, ...): add them to {S.of('Workflows')} with a Category"]


def test_a_start_in_the_second_of_the_backlog_entry_leaves_the_ticket_without_a_start():
    # created at 09:00:00 in To Do, moved to Dev in the same second by an automation (the milliseconds are dropped): no StartUtc
    quick = issue("T-3", [("2025-05-02T09:00:00", "Dev"), ("2025-05-04T00:00", "Done")], created="2025-05-02T09:00:00")
    assert fields(CLEAN + [quick]) == [" | Started in the second of the backlog entry: 1 tickets (no StartUtc, so left out of Flow Time and Flow Load)"]
    later = issue("T-4", [("2025-05-02T09:00:01", "Dev"), ("2025-05-04T00:00", "Done")], created="2025-05-02T09:00:00")
    assert fields(CLEAN + [later]) == []                                                  # one second later is a start


def test_status_changes_after_asof_say_that_asof_is_before_the_data():
    got = fields(CLEAN, as_of=day("2025-05-03T12:00"))                                    # T-1 Done 05-04, F-1 Shipped 05-05: 2 changes after
    assert got == [f" | AsOf before the data: 2 status changes after AsOf (Flow Load stops at AsOf): set AsOf in {S.of('Config')} to the time of "
                   "the extraction"]


def test_a_velocity_check_with_a_gap_and_the_controls_to_review():
    ok = VelocityCheck("I1", "Alpha", 1, 0, 0, 0, 1, 0, 0, 0)
    gap = VelocityCheck("I1", ART, 1, 1, 0, 0, 0, 0, 0, 1)
    assert fields(CLEAN, checks=[ok, gap]) == [f" | VelocityCheck gap: 1 of 2 rows (sheet {S.of('VelocityCheck')})"]
    assert fields(CLEAN, verdicts=["OK", "REVIEW", "INFO", "REVIEW"]) == [
        f" | Reconciliation: 2 of 4 controls to review (sheet {S.of('Reconciliation')})"]


def test_the_fields_come_in_the_order_of_the_vba():
    rules = Rules({k: StatusRule(r.category, False, r.is_end, False) for k, r in RULES._rows.items()})
    lost = issue("X-1", [("2025-05-02T00:00", "Odd")], team="")
    got = [f.split(":")[0] for f in fields(CLEAN + [lost], rules, as_of=day("2025-05-03T12:00"), verdicts=["REVIEW"], undated_pis=["PI#11-2027"])]
    assert got == [" | No start status", " | PIs without dates", " | Without a team", " | Statuses not in Workflows", " | AsOf before the data",
                   " | Reconciliation"]


def test_the_vba_says_the_same_words():
    vba = (Path(__file__).resolve().parents[2] / "vba" / "modHealth.bas").read_text(encoding="ascii")
    for text in H.VBA_TEXTS:
        assert f'"{text}"' in vba, text
    order = vba.split("Public Function HealthFields", 1)[1].split("Exit Function", 1)[0]
    names = ("NoStartStatus", "NoWaitStatus", "NotFeatures", "OutsideIterations", "PisWithoutDates", "WithoutTeam", "UnclassifiedStatuses", "UnknownStatuses",
             "EmptyFields", "SameSecondStarts",
             "AsOfBeforeData", "VelocityCheckGap", "ReconciliationFindings")
    assert [order.index(n + "(") for n in names] == sorted(order.index(n + "(") for n in names)


# --- the settings that match nothing (AL, 2026-11-27: the exception "À déployer" ruled nothing, WorkTypeField held the name "Activité") --------

def _rules_from_jira(exceptions=(), config=()):
    from jira2dashboard.params.model import ParamSheet
    sheets = [ParamSheet("Workflows", ("Workflow", "Status", "Category", "IsEnd", "IsFree"), tuple(exceptions)),
              ParamSheet("Config", ("Parameter", "Value", "Example", "Mandatory", "Description"), tuple(config))]
    return Rules.from_params(sheets, {"To Do": "new", "Dev": "indeterminate", "A déployer": "indeterminate", "Done": "done"}, {})


def test_a_status_that_a_setting_names_and_jira_does_not_know_is_named_with_where_it_is_typed():
    rules = _rules_from_jira([("Team", "À déployer", "Wait", None, None), (None, "dev", "Active", None, None)],
                             [("FeatureStartStatuses", "En cours, Dev", None, None, None)])
    got = [f for f in fields(CLEAN, rules) if f.startswith(H.UNKNOWN_STATUSES)]
    assert got == [f" | Statuses unknown to Jira: 2 (À déployer ({S.of('Workflows')}), En cours (FeatureStartStatuses)): type each name as Jira spells it"]


def test_a_status_known_only_by_a_history_is_known_and_jira_not_read_says_nothing():
    rules = _rules_from_jira([(None, "Review", "Wait", None, None)])                    # not in the catalog, but T-1 went through it
    assert not [f for f in fields(CLEAN, rules) if f.startswith(H.UNKNOWN_STATUSES)]
    assert H.unknown_statuses(RULES, []) == ""                                           # no catalog, no current status: nothing to compare with


def test_a_field_of_config_empty_in_every_ticket_is_named_and_a_name_is_no_id():
    config = {"WorkTypeField": "Activité", "StoryPointsField": "customfield_1", "PIField": "customfield_2", "TeamField": ""}
    header = ("Key", "WorkType", "StoryPoints", "PI", "BV", "TeamValue", "SprintState", "TargetStart", "TargetEnd")
    rows = [("A-1", None, "", "PI-1", None, None, None, None, None), ("A-2", "", None, None, None, None, None, None, None)]
    filled = H.fields_filled(config, header, rows)
    assert filled[:3] == [("WorkTypeField", "Activité", 0), ("StoryPointsField", "customfield_1", 0), ("PIField", "customfield_2", 1)]
    assert H.empty_fields(filled, 2) == (f" | Fields empty in every ticket: 2 (WorkTypeField (Activité: not a field id), StoryPointsField): type the id "
                                         f"of the field (customfield_ and digits, listed in {S.of('JiraFields')}) in {S.of('Config')}")
    assert H.empty_fields(filled, 0) == ""                                               # no ticket read: nothing to say
    assert H.empty_fields([("TeamField", "", 0)], 2) == ""                               # a field that is not set is not asked for


def test_the_fields_of_config_are_the_columns_of_issues_on_both_sides():
    from jira2dashboard.flow.formulas import ISSUES
    vba = (Path(__file__).resolve().parents[2] / "vba" / "modHealth.bas").read_text(encoding="ascii")
    assert all(column in ISSUES for _p, column in H.FIELD_COLUMNS)
    flat = re.sub(r" _\n\s+", " ", vba)                                                   # the VBA lines continued with " _"
    assert ", ".join(f'"{p}", "{c}"' for p, c in H.FIELD_COLUMNS) in flat
