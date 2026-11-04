"""The dashboard description, without Excel: every chart draws an existing column, layouts do not overlap."""
import re

from jira2dashboard.flow import dashboard_spec as d
from jira2dashboard.flow import formulas as f
from jira2dashboard.params.model import load_all

METRICS_COLUMNS = set(f.TABLES["Metrics"])


def test_one_chart_per_flow_metric():
    assert [c.name for c in d.CHARTS] == ["velocity", "flow_time", "flow_load", "flow_efficiency", "flow_distribution",
                                           "flow_predictability"]
    assert len({c.name for c in d.CHARTS}) == len(d.CHARTS)


def test_every_chart_draws_a_column_of_the_metrics_table():
    for chart in d.CHARTS:
        if chart.series == d.TEAMS:
            assert chart.column in METRICS_COLUMNS, chart.name
        else:
            assert chart.column is None and d.series_names(chart) == f.BUCKETS
            assert all(f"DistDone{b}" in METRICS_COLUMNS for b in f.BUCKETS)


def test_formulas_only_use_existing_columns_and_are_balanced():
    texts = [d.team_point(c.column, "$A6", "B$5") for c in d.CHARTS if c.column]
    texts += [d.bucket_point(b, "$A6") for b in f.BUCKETS]
    texts.append(d.note_formula("ChartData!$B$6:$J$16", "a reason"))
    for text in texts:
        assert text.startswith("=") and text.count("(") == text.count(")") and text.count('"') % 2 == 0
        for column in re.findall(r"Metrics\[(\w+)\]", text):
            assert column in METRICS_COLUMNS, column


def test_not_computable_becomes_a_gap_not_a_zero():
    assert "NA()" in d.team_point("FlowTimeMedian", "$A6", "B$5")
    assert d.bucket_point("Defect", "$A6").endswith("NA())")


def test_no_team_name_is_written_in_the_formulas():
    names = {row[2] for row in next(s for s in load_all() if s.name == "Projects").rows if row[2]}
    text = d.bucket_point("Feature", "$A6") + d.team_point("Velocity", "$A6", "B$5")
    for name in names - {"ART"}:
        assert f'"{name}"' not in text


def test_predictability_explains_why_it_has_gaps():
    spec = next(c for c in d.CHARTS if c.name == "flow_predictability")
    assert spec.reason and "PI objectives" in spec.reason
    assert "PI objectives" in d.note_formula("A1:B2", spec.reason)


def test_blocks_of_chart_data_do_not_overlap():
    periods = 11
    tops = [d.block_top(i, periods) for i in range(len(d.CHARTS))]
    assert tops[0] == 1
    for first, second in zip(tops, tops[1:]):
        assert second - first >= periods + 3          # title + series names + periods, and a blank row


def test_charts_on_the_dashboard_do_not_overlap():
    boxes = []
    for i in range(len(d.CHARTS)):
        row, column = d.chart_anchor(i)
        boxes.append((row, row + d.ROWS_PER_CHART, column, column + d.COLUMNS_PER_CHART))
    for i, a in enumerate(boxes):
        for b in boxes[i + 1:]:
            apart = a[1] < b[0] or b[1] < a[0] or a[3] < b[2] or b[3] < a[2]
            assert apart, (a, b)
    assert d.CHART_SPACING > d.ROWS_PER_CHART                # a row for the note under each chart
