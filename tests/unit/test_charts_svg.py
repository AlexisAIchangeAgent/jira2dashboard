"""The SVG charts of the Python path (pure): what they draw, gaps for "not computable", well-formed SVG."""
import datetime as dt
from xml.etree import ElementTree

import pytest

from jira2dashboard.flow.metrics import PIPredictability
from jira2dashboard.flow import dashboard_spec as D
from jira2dashboard.flow.metrics import BUCKETS, TeamIterationMetrics
from jira2dashboard.flow.model import NOT_COMPUTABLE
from jira2dashboard.standalone import charts_svg as C

SVG = "{http://www.w3.org/2000/svg}"


def shapes(svg: str, tag: str):
    return ElementTree.fromstring(svg).findall(f".//{SVG}{tag}")


def texts(svg: str) -> list[str]:
    return [t.text for t in shapes(svg, "text")]


@pytest.mark.parametrize("highest, top, step", [(0, 1.0, 0.2), (7, 8, 2), (10, 10, 2), (23, 25, 5), (0.83, 1.0, 0.2),
                                                (130, 150, 50), (4.2, 5, 1)])
def test_the_axis_ends_on_a_round_number(highest, top, step):
    assert C.nice_top(highest) == (pytest.approx(top), pytest.approx(step))


def test_columns_one_bar_per_figure_and_none_for_a_gap():
    svg = C.draw("Velocity", ["I1", "I2"], {"ART": [3, None], "Alpha": [0, 5]}, C.COLUMNS)
    bars = [r for r in shapes(svg, "rect") if r.get("fill") in C.PALETTE]
    assert len(bars) == 3 + 2                         # 3 figures (a 0 is a bar of no height) + 2 legend boxes
    assert "Velocity" in texts(svg) and "ART" in texts(svg) and "Alpha" in texts(svg)


def test_a_line_breaks_at_a_gap_and_keeps_its_markers():
    svg = C.draw("Flow Time", ["I1", "I2", "I3", "I4"], {"ART": [1.5, 2.0, None, 4.0]}, C.LINES)
    assert len(shapes(svg, "polyline")) == 1 and len(shapes(svg, "circle")) == 3
    assert len(shapes(C.draw("Load", ["d1", "d2"], {"A": [1, 2]}, C.LINE_ONLY), "circle")) == 0


def test_shares_stack_to_the_full_height_in_percent():
    svg = C.draw("Distribution", ["I1"], {"Feature": [0.5], "Enabler": [0.25], "Defect": [0.25], "Autre": [0.0],
                                          "Unmapped": [None]}, C.STACKED)
    assert "100%" in texts(svg) and "50%" not in texts(svg) and "40%" in texts(svg)
    assert len([r for r in shapes(svg, "rect") if r.get("fill") in C.PALETTE]) == 3 + 5


def test_nothing_to_draw_says_not_computable_with_the_band_and_the_note():
    series = {"ART": [None, None]}
    svg = C.draw("Flow Predictability", ["I1", "I2"], series, C.LINES, True, (0.8, 1.0),
                 C.gaps_note(series, "it needs the PI objectives, which are not in Jira"))
    assert NOT_COMPUTABLE in texts(svg)
    assert "2 of 2 points are not computable (gaps) - it needs the PI objectives, which are not in Jira" in texts(svg)
    assert any(r.get("fill") == "#e3f2e1" for r in shapes(svg, "rect"))
    assert C.gaps_note({"A": [1, 2]}, "why") == ""


def test_text_is_escaped():
    svg = C.draw("A <b> & c", ["I<1>"], {"R&D": [1]}, C.COLUMNS)
    assert "A <b> & c" in texts(svg) and "R&D" in texts(svg)


def metrics(iteration, team, velocity, done=None, median=NOT_COMPUTABLE):
    shares = done or {b: (1.0 if b == "Feature" else 0.0) for b in BUCKETS}
    return TeamIterationMetrics(iteration, team, velocity, velocity, 0, median, median, 1.0, 1, 0.5,
                             shares if velocity else NOT_COMPUTABLE, NOT_COMPUTABLE, NOT_COMPUTABLE)


def test_all_teams_weights_each_team_by_its_velocity():
    defects = {b: (1.0 if b == "Defect" else 0.0) for b in BUCKETS}
    rows = [metrics("I1", "ART", 1), metrics("I1", "Alpha", 3, defects), metrics("I2", "ART", 0)]
    shares = C.distribution(rows, ["I1", "I2"], D.ALL_TEAMS)
    assert shares["Feature"] == [0.25, None] and shares["Defect"] == [0.75, None]
    assert C.distribution(rows, ["I1"], "Alpha")["Defect"] == [1.0]


def test_the_eight_charts_of_the_dashboard():
    teams, iterations = ["ART", "Alpha"], ["I1", "I2"]
    rows = [metrics(p, t, v, median=2.5) for p in iterations for t, v in zip(teams, (2, 0))]
    days = [dt.datetime(2026, 10, 1) + dt.timedelta(days=n) for n in range(5)]
    loads = {"ART": {d: 1 for d in days[:3]}, "Alpha": {d: 0 for d in days[:3]}}
    pictures = C.chart_pictures(rows, teams, iterations, loads, days, "ART",
                                [PIPredictability("PI#01-2026", 4, 1, 0.25), PIPredictability("PI#02-2026", 0, 0, NOT_COMPUTABLE)])
    assert [p.name for p in pictures] == [spec.name for spec in D.CHARTS]
    feature = next(p for p in pictures if p.name == "feature_predictability")
    assert "PI#01-2026" in texts(feature.svg) and "PI#02-2026" in texts(feature.svg)        # one bar per PI, a gap for PI 2
    assert C.chart_pictures(rows, teams, iterations, loads, days, "ART")[-2].name == "feature_predictability"
    for picture in pictures:
        ElementTree.fromstring(picture.svg)
    by_name = {p.name: p for p in pictures}
    assert "Flow Distribution - items done by work type (ART)" in texts(by_name["flow_distribution"].svg)
    assert "4 of 10 points are not computable (gaps) - the day is after the as-of date" in texts(
        by_name["flow_load_daily"].svg)
    assert any(t == "10-03" for t in texts(by_name["flow_load_daily"].svg))


def test_counts_get_whole_gridlines():
    labels = texts(C.draw("Velocity", ["I1"], {"ART": [2]}, C.COLUMNS))
    assert "1" in labels and "2" in labels and "0.5" not in labels and "1.5" not in labels
    assert "0.5" in texts(C.draw("Flow Time", ["I1"], {"ART": [1.6]}, C.LINES))
