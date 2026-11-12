"""Feature Predictability, the variant of Flow Predictability AL chose (2026-11-12): for each PI, the Features done before
the end of the PI over the Features planned in it (their PI field is that PI). A Feature cancelled stays planned, not done;
one done after the end of its PI is not done for that PI. ART only: one ratio per PI. Hand-written cases, results by hand.
"""
import dataclasses

from conftest import FLOW_AS_OF, at, flow_issue
from jira2dashboard.flow.checks import CheckSettings
from jira2dashboard.flow.itemflow import item_flow
from jira2dashboard.flow.metrics import PIPredictability, feature_predictability
from jira2dashboard.flow.model import NOT_COMPUTABLE, Rules

FEATURE_FLOW = [("10-02", "En analyse"), ("10-03", "Prêt"), ("10-04", "Planifié"), ("10-05", "En développement"),
                ("10-08", "En test / recette"), ("10-09", "À déployer")]
SETTINGS = CheckSettings(feature_types=frozenset({"Feature"}),
                         pis={"PI#01-2026": (at("10-01"), at("10-20")), "PI#02-2026": (at("10-20"), at("11-30")),
                              "PI#03-2026": (at("12-01"), at("12-31"))})


def feature(key, pi, moves, type_="Feature"):
    return dataclasses.replace(flow_issue(key, "ART", type_, "Feature", "09-30", moves, workflow="Feature"), pi=pi)


def cases():
    return [
        feature("F-1", "PI#01-2026", FEATURE_FLOW + [("10-10", "Terminé")]),           # done before the end of PI 1
        feature("F-2", "PI#01-2026", FEATURE_FLOW + [("10-25", "Terminé")]),           # done, but after the end of PI 1
        feature("F-3", "PI#01-2026", FEATURE_FLOW[:3] + [("10-06", "Annulé")]),        # cancelled: planned, not done
        feature("F-4", "PI#01-2026", FEATURE_FLOW[:4]),                                # still in progress
        feature("F-5", "PI#02-2026", FEATURE_FLOW + [("10-21", "Terminé")]),           # PI 2: done
        feature("S-1", "PI#01-2026", FEATURE_FLOW + [("10-10", "Terminé")], "Story"),  # not a Feature type: left out
        feature("F-6", None, FEATURE_FLOW + [("10-10", "Terminé")]),                   # no PI: in no PI
    ]


def test_each_pi_gives_its_features_done_before_its_end_over_its_features_planned():
    issues = cases()
    items = [item_flow(issue, FLOW_AS_OF, Rules.from_params()) for issue in issues]
    assert feature_predictability(issues, items, SETTINGS) == [
        PIPredictability("PI#01-2026", planned=4, done=1, ratio=0.25),       # F-1 of F-1..F-4 (F-2 late, F-3 cancelled, F-4 open)
        PIPredictability("PI#02-2026", planned=1, done=1, ratio=1.0),
        PIPredictability("PI#03-2026", planned=0, done=0, ratio=NOT_COMPUTABLE),        # nothing planned: a gap, never 0
    ]


def test_without_any_pi_in_the_sheet_there_is_nothing_to_show():
    issues = cases()
    items = [item_flow(issue, FLOW_AS_OF, Rules.from_params()) for issue in issues]
    assert feature_predictability(issues, items, CheckSettings(feature_types=frozenset({"Feature"}))) == []


def test_pimetrics_rows_are_the_table_of_the_workbook_one_row_per_pi_or_the_row_no_pi():
    from types import SimpleNamespace
    from jira2dashboard.flow import dashboard_spec as D
    from jira2dashboard.flow import formulas as F
    from jira2dashboard.standalone import results as R
    start, end = SETTINGS.pis["PI#01-2026"]
    s = SimpleNamespace(checks=SETTINGS)
    found = SimpleNamespace(predictability=[PIPredictability("PI#01-2026", 4, 1, 0.25)])
    rows = R.pimetrics_rows(found, s)
    assert rows == [("PI#01-2026", start, end, 4, 1, 0.25)] and len(rows[0]) == len(F.TABLES["PIMetrics"])
    assert R.pimetrics_rows(SimpleNamespace(predictability=[]), s) == [(D.NO_PI, None, None, 0, None, NOT_COMPUTABLE)]
