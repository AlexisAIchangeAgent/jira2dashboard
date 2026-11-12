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


def test_a_pi_that_features_carry_but_the_sheet_lacks_is_shown_after_them_not_computable():
    """AL (2026-11-12): a PI missing from the sheet PIs has no end date, so its ratio is not computable, but the RTE sees it.
    By year, then number; a malformed PI (DC17, DC29), another type, or a PI the sheet holds add no row."""
    issues = cases() + [
        feature("F-7", "PI#01-2027", FEATURE_FLOW),
        feature("F-8", "PI#04-2026", FEATURE_FLOW + [("10-10", "Terminé")]),         # done, but no end date to compare with
        feature("F-9", "PI#04-2026", FEATURE_FLOW[:2]),
        feature("F-10", "PI 4", FEATURE_FLOW),                                       # malformed: DC29, not a PI
        feature("F-11", "pi#05-2026", FEATURE_FLOW),                                 # PI# in capitals only, as DC17
        feature("F-12", "PI#02-2026", FEATURE_FLOW),                                 # in the sheet: counted there
        feature("S-2", "PI#06-2026", FEATURE_FLOW, "Story"),                         # not a Feature type
    ]
    items = [item_flow(issue, FLOW_AS_OF, Rules.from_params()) for issue in issues]
    rows = feature_predictability(issues, items, SETTINGS)
    assert [p.pi for p in rows[:3]] == list(SETTINGS.pis) and rows[1].planned == 2
    assert rows[3:] == [PIPredictability("PI#04-2026", planned=2, done=None, ratio=NOT_COMPUTABLE),
                        PIPredictability("PI#01-2027", planned=1, done=None, ratio=NOT_COMPUTABLE)]


def test_without_any_pi_in_the_sheet_the_pis_of_the_features_are_listed():
    issues = cases()
    items = [item_flow(issue, FLOW_AS_OF, Rules.from_params()) for issue in issues]
    no_sheet = CheckSettings(feature_types=frozenset({"Feature"}))
    assert feature_predictability(issues, items, no_sheet) == [
        PIPredictability("PI#01-2026", planned=4, done=None, ratio=NOT_COMPUTABLE),
        PIPredictability("PI#02-2026", planned=1, done=None, ratio=NOT_COMPUTABLE)]
    no_pi = [dataclasses.replace(issue, pi=None) for issue in issues]
    assert feature_predictability(no_pi, items, no_sheet) == []                 # nothing to show: the row NO_PI


def test_pimetrics_rows_are_the_table_of_the_workbook_one_row_per_pi_or_the_row_no_pi():
    from types import SimpleNamespace
    from jira2dashboard.flow import dashboard_spec as D
    from jira2dashboard.flow import formulas as F
    from jira2dashboard.standalone import results as R
    start, end = SETTINGS.pis["PI#01-2026"]
    s = SimpleNamespace(checks=SETTINGS)
    found = SimpleNamespace(predictability=[PIPredictability("PI#01-2026", 4, 1, 0.25),
                                            PIPredictability("PI#04-2026", 2, None, NOT_COMPUTABLE)])     # not in the sheet
    rows = R.pimetrics_rows(found, s)
    assert rows == [("PI#01-2026", start, end, 4, 1, 0.25), ("PI#04-2026", None, None, 2, None, NOT_COMPUTABLE)]
    assert len(rows[0]) == len(F.TABLES["PIMetrics"])
    assert R.pimetrics_rows(SimpleNamespace(predictability=[]), s) == [(D.NO_PI, None, None, 0, None, NOT_COMPUTABLE)]
