"""The report expected / obtained of the scenario art-metrics (tests/art_report.py): pure, tested without Excel."""
import art_report as R
from art_oracle import (ART_ROWS, CHAIN, DONE_A, DONE_B, FLOW_EFFICIENCY, FLOW_TIME_DAYS, PI_ITEMS, PIS, PLANNED, STATUS_AVERAGE_DAYS,
                        STATUS_SHARE, TEAM_ROWS, VELOCITY)
from jira2dashboard.flow.model import NOT_COMPUTABLE


def perfect_tables(mode="A"):
    """The tables a correct workbook would hold, written from the oracle."""
    done = DONE_A if mode == "A" else DONE_B
    pimetrics = [{"PI": pi, "FeaturesPlanned": PLANNED[pi], "FeaturesDone": done[pi], "FeaturePredictability": done[pi] / PLANNED[pi]} for pi in PIS]
    features = [{"Iteration-label": label, "Team": "ART", **dict(zip(R.ART_COLUMNS, values))} for label, values in ART_ROWS.items()]
    features += [{"Iteration-label": label, "Team": team, **dict(zip(R.TEAM_COLUMNS, values))}
                 for team, rows in TEAM_ROWS.items() for label, values in rows.items()]
    metrics = []
    for label in ART_ROWS:
        released = VELOCITY[label]
        metrics.append({"Iteration-label": label, "Team": "ART", "Velocity": released,
                        "FlowTimeMedian": FLOW_TIME_DAYS if released else NOT_COMPUTABLE,
                        "FlowTimeP85": FLOW_TIME_DAYS if released else NOT_COMPUTABLE,
                        "FlowEfficiency": FLOW_EFFICIENCY["PI-" + label[3]] if released else NOT_COMPUTABLE})
    status_time = [{"PI": pi, "Status": s, "Items": PI_ITEMS[pi], "AverageDays": STATUS_AVERAGE_DAYS[pi][s], "Share": STATUS_SHARE[pi][s]}
                   for pi in PIS for s in CHAIN]
    checks = {table: [{"Iteration-label": label, "Team": "ART", **dict(zip(columns, values))} for label, values in expected.items()]
              for table, _tells, columns, expected, _percent in R.CHECK_TABLES}
    return {"PIMetrics": pimetrics, "IterationFeatures": features, "Metrics": metrics, "PIStatusTime": status_time, **checks}


def rows_of(found):
    return [r for _title, rows in found for r in rows]


def test_a_workbook_that_holds_the_hand_made_numbers_is_all_ok_in_both_modes():
    for mode in ("A", "B"):
        found = R.sections(perfect_tables(mode), mode)
        assert [title.split(" (")[0] for title, _rows in found] == [
            "Feature Predictability of each PI", "Features of the ART per iteration", "Features of each team, PI-3 and PI-4",
            "Velocity of the ART per iteration", "Flow Time median / p85", "Cycle time by status, per PI",
            "FlowTimeCheck of the ART per iteration: why the Flow Time is what it is",
            "FlowEfficiencyCheck of the ART per iteration: why the Flow Efficiency is what it is",
            "FlowLoadCheck of the ART per iteration: why the Flow Load is what it is",
            "FlowDistCheck of the ART per iteration: what the Flow Distribution counts"]
        assert all(ok for _l, _e, _o, ok in rows_of(found)), [r for r in rows_of(found) if not r[3]][:3]
        assert len(rows_of(found)) == 4 + 24 + 36 + 24 + 24 + 28 + 4 * 24


def test_the_wrong_mode_is_a_difference_not_a_crash():
    found = R.sections(perfect_tables("A"), "B")                        # the workbook is in mode A, the report expects B: only PI-4 differs
    assert [label for label, _e, _o, ok in rows_of(found) if not ok] == ["PI-4: planned / done (mode B) / predictability"]


def test_a_figure_that_differs_is_marked_with_both_values_shown():
    tables = perfect_tables()
    next(r for r in tables["IterationFeatures"] if r["Team"] == "ART" and r["Iteration-label"] == "IT-3-4")["Released"] = 1
    tables["PIMetrics"][2]["FeaturePredictability"] = 0.4
    bad = {label: (expected, obtained) for label, expected, obtained, ok in rows_of(R.sections(tables)) if not ok}
    assert bad["IT-3-4"] == ("3 / 0 / 0 / 2 / 0 / 1", "3 / 0 / 0 / 1 / 0 / 1")
    assert bad["PI-3: planned / done (mode A) / predictability"] == ("6 / 3 / 50.00 %", "6 / 3 / 40.00 %")


def test_a_missing_row_or_value_is_a_difference():
    tables = perfect_tables()
    tables["Metrics"] = [r for r in tables["Metrics"] if r["Iteration-label"] != "IT-1-1"]
    tables["PIMetrics"][0]["FeaturesDone"] = None
    bad = [label for label, _e, _o, ok in rows_of(R.sections(tables)) if not ok]
    assert "IT-1-1" in bad and "PI-1: planned / done (mode A) / predictability" in bad


def test_the_percentages_and_the_flow_figures_are_shown_as_a_reader_expects():
    predictability, per_iteration, _teams, _velocity, flow, cycle, *checks = [rows for _title, rows in R.sections(perfect_tables())]
    assert cycle[0][1:3] == ("6 / 1 / 12.50 %", "6 / 1 / 12.50 %") and cycle[4][0] == "PI-1 testing" and cycle[4][1] == "6 / 2 / 25.00 %"
    assert per_iteration[0][1:3] == ("6 / 0 / 0 / 1 / 0 / 5", "6 / 0 / 0 / 1 / 0 / 5")
    assert flow[0][1:3] == ("8 / 8 / 62.50 %", "8 / 8 / 62.50 %")                                   # IT-1-1 releases
    assert flow[11][0] == "IT-2-6" and flow[11][1] == f"{NOT_COMPUTABLE} / {NOT_COMPUTABLE}"         # IT-2-6 releases nothing
    assert predictability[0][1] == "6 / 6 / 100.00 %"
    time_check, efficiency_check, _load, distribution = checks
    assert time_check[0][1] == "1 / 0 / 0 / 0 / 0 / 1 / 0 / -" and efficiency_check[0][1] == "1 / 0 / 5 / 3 / 3 / 62.50 % / 0 / -"   # IT-1-1: one release
    assert distribution[11][1] == "0 / 0 / 0 / 0 / 0 / 0 / 0 / no item done in the iteration / -"                              # IT-2-6: none


def test_the_html_says_the_verdict_and_opens_the_sections_that_differ_with_the_differences_first():
    ok_page = R.render(R.sections(perfect_tables()))
    assert "All 236 figures are the ones worked out by hand." in ok_page and "<details class=\"rep\" open" not in ok_page
    tables = perfect_tables()
    tables["Metrics"][3]["Velocity"] = 9
    page = R.render(R.sections(tables))
    assert "1 of 236 figures differ" in page and '<details class="rep" open>' in page
    velocity = page.split("Velocity of the ART")[1].split("</details>")[0]
    assert velocity.index('class="diff"') < velocity.index('class="ok"')               # the line that differs comes first


def test_an_accounting_table_that_differs_or_is_missing_shows_and_an_empty_text_read_as_none_does_not():
    tables = perfect_tables()
    next(r for r in tables["FlowTimeCheck"] if r["Iteration-label"] == "IT-1-1")["Gap"] = 1
    for r in tables["FlowLoadCheck"]:
        r["Reason"] = None                                                   # the Python path reads "" back as None
    del tables["FlowDistCheck"]
    bad = [(title.split(" of")[0], label) for title, rows in R.sections(tables) for label, _e, _o, ok in rows if not ok]
    assert ("FlowTimeCheck", "IT-1-1") in bad and not any(t == "FlowLoadCheck" for t, _l in bad)
    assert sum(1 for t, _l in bad if t == "FlowDistCheck") == 24 and len(bad) == 25


def test_the_html_escapes_what_it_shows():
    page = R.render([("a <b> title", [("x<y", "1", "2", False)])])
    assert "a &lt;b&gt; title" in page and "x&lt;y" in page and "<b>" not in page


def test_same_compares_numbers_with_a_margin_and_texts_exactly():
    assert R.same(0.625, 0.625 + 1e-12) and not R.same(1, 2) and not R.same(1, "1") and R.same("a", "a") and not R.same(1, None)
