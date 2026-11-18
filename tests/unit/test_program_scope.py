"""The ART counts the Features only (AL, 2026-11-18): a ticket of the Program project whose type is not in Config.FeatureTypes is reported under
the team "ART (other types)", so no metric of the ART counts it, while it stays in ItemFlow (findable by its key) and in the data checks.

Why: on the client PC the Program project TSF holds the Features ("Fonctionnalité") and their sub-tasks; the Velocity of the ART counted both (60 for 8 Features).
"""
import datetime as dt

from jira2dashboard.flow.model import ART, OTHER_TYPES, FlowIssue, program_team
from jira2dashboard.mockjira import flow_adapter
from jira2dashboard.mockjira.scenarios import get
from jira2dashboard.standalone import mapping
from jira2dashboard.standalone.results import fill_teams

FEATURES = frozenset({"Fonctionnalité"})
PROJECTS = [("TSF", "Program", ""), ("TEAM", "Team", "")]
DAY = dt.datetime(2026, 10, 1)


def row(key, project, type_):
    """A row of the sheet Issues (the 22 columns that mapping.flow_issue unpacks), nothing else filled."""
    return (key, "1", project, type_, None, DAY, "Nouveau", None, "new", None, "", None, None, None, None, "", None, None, 0, None, None, 0)


def test_the_team_of_a_program_ticket_is_art_for_a_feature_and_the_other_team_for_the_rest():
    assert OTHER_TYPES == "ART (other types)"
    assert program_team("Fonctionnalité", FEATURES) == ART
    assert program_team("Sous-tâche", FEATURES) == OTHER_TYPES
    assert program_team("Sous-tâche", None) == ART                      # None: no filter (the callers that do not know the parameter)
    assert program_team("Fonctionnalité", frozenset()) == OTHER_TYPES   # an empty FeatureTypes: no Feature at all


def test_the_extraction_of_the_python_path_gives_the_other_team_to_the_other_types_of_the_program_only():
    got = {key: mapping.flow_issue(row(key, project, type_), [], PROJECTS, (), (), feature_types=FEATURES).team
           for key, project, type_ in (("T-1", "TSF", "Fonctionnalité"), ("T-2", "TSF", "Sous-tâche"), ("U-1", "TEAM", "Sous-tâche"))}
    assert got == {"T-1": ART, "T-2": OTHER_TYPES, "U-1": "TEAM"}     # a Team project is untouched: the Team project key is its team
    assert mapping.flow_issue(row("T-2", "TSF", "Sous-tâche"), [], PROJECTS, (), ()).team == ART     # no feature_types: as before


def test_the_adapter_of_the_fake_jira_does_the_same():
    ds = get("art-metrics").dataset
    default = {i.key: i.team for i in flow_adapter.to_flow_issues(ds)}                                 # default: FeatureTypes of the sheet Config (Feature)
    assert {k: t for k, t in default.items() if k.split("-")[1] in ("26", "27", "28")} == {"TTT-26": OTHER_TYPES, "TTT-27": OTHER_TYPES, "TTT-28": OTHER_TYPES}
    assert {t for k, t in default.items() if k.split("-")[1] not in ("26", "27", "28")} == {ART}       # the 25 Features
    assert {i.team for i in flow_adapter.to_flow_issues(ds, feature_types=frozenset({"Epic"}))} == {OTHER_TYPES}
    assert {i.team for i in flow_adapter.to_flow_issues(ds, feature_types=frozenset({"Feature", "Sub-task"}))} == {ART}


def test_fill_rates_of_the_python_path_do_not_list_the_other_team():
    issues = [FlowIssue(key="T-1", team=ART, type="Fonctionnalité", work_type=None, workflow="Feature", created=DAY, status="Nouveau"),
              FlowIssue(key="T-2", team=OTHER_TYPES, type="Sous-tâche", work_type=None, workflow="Feature", created=DAY, status="Nouveau"),
              FlowIssue(key="U-1", team="TEAM", type="Story", work_type=None, workflow="Team", created=DAY, status="Nouveau")]
    assert fill_teams(issues) == [ART, "TEAM"]                          # the workbook's FillRates keys: ART, then the Team projects
