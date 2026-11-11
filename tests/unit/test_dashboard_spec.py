"""The dashboard description, without Excel: every chart draws an existing column, layouts do not overlap."""
import re

from jira2dashboard.flow import dashboard_spec as d
from jira2dashboard.flow import formulas as f
from jira2dashboard.params.model import load_all

METRICS_COLUMNS = set(f.TABLES["Metrics"])


def test_one_chart_per_flow_metric():
    assert [c.name for c in d.CHARTS] == ["velocity", "flow_time", "flow_time_p85", "flow_load", "flow_efficiency",
                                           "flow_distribution", "flow_predictability", "flow_load_daily"]
    assert len({c.name for c in d.CHARTS}) == len(d.CHARTS)


def test_every_chart_draws_a_column_of_the_metrics_table():
    for chart in d.CHARTS:
        if chart.axis == d.DAY_AXIS:
            assert chart.column is None and {"Date", "Team", "Load"} <= set(f.TABLES["DailyLoad"])
        elif chart.series == d.TEAMS:
            assert chart.column in METRICS_COLUMNS, chart.name
        else:
            assert chart.column is None and d.series_names(chart) == f.BUCKETS
            assert all(f"DistDone{b}" in METRICS_COLUMNS for b in f.BUCKETS)


def test_formulas_only_use_existing_columns_and_are_balanced():
    texts = [d.team_point(c.column, "$A6", "B$5") for c in d.CHARTS if c.column]
    texts += [d.bucket_point(b, "$A6") for b in f.BUCKETS] + [d.day_point("$A6", "B$5")]
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
    iterations = 11
    tops = [d.block_top(i, iterations) for i in range(len(d.CHARTS))]
    assert tops[0] == 1
    for first, second in zip(tops, tops[1:]):
        assert second - first >= iterations + 3          # title + series names + iterations, and a blank row


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


def test_the_daily_curve_is_last_and_has_a_gap_for_days_without_snapshot():
    daily = d.CHARTS[-1]
    assert daily.axis == d.DAY_AXIS and all(c.axis == d.ITERATION_AXIS for c in d.CHARTS[:-1])      # its block is the longest
    text = d.day_point("$A6", "B$5")
    assert "NA()" in text and "DailyLoad[Load]" in text and text.count("(") == text.count(")")


def test_the_predictability_band_is_a_proposal_text_the_vba_can_read():
    spec = next(c for c in d.CHARTS if c.name == "flow_predictability")
    assert spec.band == (0.8, 1.0) and d.band_text(spec) == "0.8;1.0"
    assert all(d.band_text(c) == "" for c in d.CHARTS if c is not spec)


def test_the_vba_names_the_same_sheet_of_charts_and_the_same_status_rows():
    """modFlowBuild draws the charts, modActions names them in the status Done and writes the status rows: their constants
    are those of dashboard_spec (the build makes the sheet Charts and formats the rows)."""
    from jira2dashboard.testbench import run_vba_tests as rvt
    sources = rvt.module_sources(rvt.ORDER)
    for module in ("modFlowBuild", "modActions"):
        assert f'Private Const CHARTS_SHEET As String = "{d.CHARTS_SHEET}"' in sources[module], module
    assert f"Private Const STATUS_ROWS As Long = {d.STATUS_ROWS}" in sources["modActions"]
