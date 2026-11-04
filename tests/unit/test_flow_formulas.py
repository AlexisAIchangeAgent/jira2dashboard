"""The text of the Excel formulas, checked without Excel: references resolve, no mapping is written in them."""
import re
import unicodedata

import pytest

from jira2dashboard.flow import formulas as f
from jira2dashboard.flow import metrics, model
from jira2dashboard.params.model import load_all

OWN_TABLE = {}
for name, text in f.INTERVAL_HELPERS:
    OWN_TABLE[("StatusIntervals", name)] = text
for name, text in f.ITEMFLOW_COLUMNS:
    OWN_TABLE[("ItemFlow", name)] = text
for name, text in f.METRICS_COLUMNS:
    OWN_TABLE[("Metrics", name)] = text
OWN_TABLE[("DailyLoad", "Load")] = f.DAILYLOAD_LOAD

ALL = list(OWN_TABLE.items())
TABLE_REF = re.compile(r"\b(Issues|StatusIntervals|ItemFlow|DailyLoad|Metrics)\[(#Headers|\w+)\]")
ROW_REF = re.compile(r"\[@(\w+)\]")


def ascii_fold(text: str) -> str:
    return unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")


@pytest.mark.parametrize("key,formula", ALL, ids=[f"{t}.{c}" for (t, c), _ in ALL])
def test_every_reference_points_to_an_existing_column(key, formula):
    table, _column = key
    for ref_table, column in TABLE_REF.findall(formula):
        assert column == "#Headers" or column in f.TABLES[ref_table], f"{ref_table}[{column}]"
    for column in ROW_REF.findall(formula):
        assert column in f.TABLES[table], f"[@{column}] in {table}"


@pytest.mark.parametrize("key,formula", ALL, ids=[f"{t}.{c}" for (t, c), _ in ALL])
def test_formulas_are_balanced(key, formula):
    assert formula.startswith("=")
    assert formula.count("(") == formula.count(")")
    assert formula.count('"') % 2 == 0
    assert formula.count("[") == formula.count("]")


def test_a_column_only_uses_columns_computed_before_it_or_raw_data():
    """No formula of ItemFlow reads another ItemFlow column that is defined after it, except through its own row
    (Excel would accept it, but the order of the table then hides the dependencies)."""
    names = [n for n, _ in f.ITEMFLOW_COLUMNS]
    for index, (name, formula) in enumerate(f.ITEMFLOW_COLUMNS):
        for used in ROW_REF.findall(formula):
            assert names.index(used) < index, f"{name} reads [@{used}] defined later"


def test_table_column_names_are_unique():
    for table, columns in f.TABLES.items():
        assert len(columns) == len(set(columns)), table


def test_no_status_name_and_no_mapping_is_written_in_a_formula():
    sheets = {s.name: s for s in load_all()}
    statuses = {ascii_fold(row[2]) for row in sheets["Workflows"].rows}
    team_names = {row[2] for row in sheets["Projects"].rows if row[2]} | {row[0] for row in sheets["Projects"].rows}
    for (table, column), formula in ALL + [(("Metrics", "AsOf"), f.AS_OF_CELL), (("Metrics", "Mode"), f.MODE_CELL)]:
        folded = ascii_fold(formula)
        for status in statuses:
            assert f'"{status}"' not in folded, f"{status} in {table}.{column}"
        for team in team_names - {"ART"}:
            assert f'"{team}"' not in formula, f"{team} in {table}.{column}"


def test_the_only_names_used_are_the_two_control_cells():
    names = {"AsOfUtc", "cfg_FlowLoadMode"}
    pattern = re.compile(r"\b(AsOfUtc|cfg_\w+)\b")
    used = {m for _k, formula in ALL for m in pattern.findall(formula)}
    assert used == names


def test_the_vocabulary_is_the_one_of_the_python_reference():
    assert f.BUCKETS == metrics.BUCKETS and f.WORK_TYPES == model.WORK_TYPES
    assert f.NOT_COMPUTABLE == model.NOT_COMPUTABLE
    assert set(f.CATEGORIES_IN_PROCESS) == {model.ACTIVE, model.WAIT, model.UNCLASSIFIED}
    assert f.INTERVALS_RAW[:2] == ("Key", "Workflow")


def test_the_metrics_cover_every_metric_and_bucket():
    columns = [n for n, _ in f.METRICS_COLUMNS]
    for expected in ("Velocity", "FlowTimeMedian", "FlowTimeP85", "FlowLoadAverage", "FlowLoadEnd", "FlowEfficiency",
                     "FlowPredictability"):
        assert expected in columns
    for prefix in ("DistDone", "WipCount", "DistWip"):
        assert [c for c in columns if c.startswith(prefix)] == [prefix + b for b in f.BUCKETS]


def test_dates_are_compared_with_a_half_second_margin():
    assert f.HALF in dict(f.ITEMFLOW_COLUMNS)["StartUtc"]
    checked = 0
    for table, formulas in (("ItemFlow", f.ITEMFLOW_COLUMNS), ("Metrics", f.METRICS_COLUMNS),
                            ("StatusIntervals", f.INTERVAL_HELPERS)):
        for name, formula in formulas:
            for bound in re.findall(r'"[<>]=?"&(\([^"]*)', formula):       # every bound built from a date
                assert "0.5/86400" in bound, (table, name, bound)
                checked += 1
    assert checked > 10
    assert all("0.5/86400" in b for b in re.findall(r'"[<>]=?"&(\([^"]*)', f.DAILYLOAD_LOAD))


def test_flow_predictability_is_always_not_computable():
    assert dict(f.METRICS_COLUMNS)["FlowPredictability"] == '="not computable"'
