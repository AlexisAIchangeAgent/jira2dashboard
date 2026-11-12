"""The plan written into the workbook (FlowFormulas, ChartSpec), without Excel."""
import re

from jira2dashboard.flow import blueprint as B
from jira2dashboard.flow import dashboard_spec as D
from jira2dashboard.flow import formulas as F
from jira2dashboard.params.model import load_all


def test_every_table_column_of_the_formulas_is_in_the_plan_in_order():
    rows = B.formula_rows()
    for table, columns in F.TABLES.items():
        assert [c for t, c, _f in rows if t == table] == list(columns), table


def test_values_columns_come_first_and_formula_columns_after():
    """modFlowBuild counts the raw columns as the columns without formula: they must be the first ones."""
    for table in ("Issues", "StatusIntervals", "ItemFlow", "DailyLoad", "Metrics"):
        flags = [bool(f) for t, _c, f in B.formula_rows() if t == table]
        assert flags == sorted(flags), table                      # False... then True...
    assert not any(f for t, _c, f in B.formula_rows() if t == "Issues")
    assert all(f for t, _c, f in B.formula_rows() if t == "ItemFlow")


def test_the_control_cells_are_named_as_the_formulas_expect():
    names = [c for t, c, _f in B.formula_rows() if t == "Controls"]
    assert tuple(names) == B.CONTROL_NAMES
    uses = set(re.findall(r"\b(AsOfUtc|cfg_\w+)\b", " ".join(f for _t, _c, f in B.formula_rows())))
    assert uses == set(B.CONTROL_NAMES)


def test_a_formula_fits_in_a_cell():
    assert max(len(f) for _t, _c, f in B.formula_rows()) < 32000
    assert max(len(r[5]) for r in B.chart_rows()) < 32000


def test_one_chart_row_per_chart_with_the_tokens_the_vba_replaces():
    rows = B.chart_rows()
    assert [r[0] for r in rows] == [c.name for c in D.CHARTS]
    for name, title, kind, axis, series, point, note, row, column, height, width, day_or_iteration, band in rows:
        assert "{iteration}" in point and "{range}" in note
        if day_or_iteration == D.PI_AXIS:                    # one fixed series (ART), its point names neither
            assert series == "ART" and "{series}" not in point and "{name}" not in point, name
        else:
            assert ("{series}" in point) == bool(series) and ("{name}" in point) == (not series), name
        assert kind in (D.XL_COLUMN_CLUSTERED, D.XL_LINE_MARKERS, D.XL_COLUMN_STACKED_100, D.XL_LINE)
        assert day_or_iteration in (D.ITERATION_AXIS, D.DAY_AXIS, D.PI_AXIS)
        assert row >= 4 and column >= 1 and height > 0 and width > 0
    assert len(B.CHARTS_HEADER) == len(rows[0])


def test_the_dashboard_settings_are_in_the_plan():
    settings = {c: f for t, c, f in B.formula_rows() if t == "Dashboard"}
    assert settings == {"ProgramTeam": "ART", "AllTeams": D.ALL_TEAMS, "SelectorCell": D.SELECTOR_CELL, "NoPI": D.NO_PI}


def test_teams_are_art_then_the_team_projects_each_once():
    projects = [("TTT", "Program", ""), ("AAA", "Team", "Alpha"), ("BBB", "Team", " Beta "), ("CCC", "team", "Alpha"),
                ("DDD", "Team", "")]
    assert B.team_names(projects) == ["ART", "Alpha", "Beta"]
    sheet = next(s for s in load_all() if s.name == "Projects")
    teams = B.team_names([tuple(r) for r in sheet.rows])
    assert teams[0] == "ART" and len(teams) == 1 + len(sheet.rows) - 1       # the Program is one row, the others teams
