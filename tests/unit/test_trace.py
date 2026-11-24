"""The trace of a ticket (flow/trace.py): in which iteration and PI it was done, in which account of the population it sits, why it is not in
the Flow Time, how far its end is from an iteration boundary (AL, 2026-11-18: "c'est de la comptabilité", every ticket in exactly one account).

Every expected figure is typed by hand from the small tickets below (iterations I1 = 05-01 .. 05-11 and I2 = 05-11 .. 05-21, End excluded; PI P1 =
05-01 .. 06-01), with the Feature workflow of params/workflows.md (Nouveau, En analyse = start, ..., Terminé = end, Annulé = removed).
"""
import datetime as dt

import pytest

from jira2dashboard.flow.itemflow import item_flow
from jira2dashboard.flow.model import OTHER_TYPES, Change, FlowIssue, Iteration, Rules, StatusRule
from jira2dashboard.flow.trace import (ACCOUNT_COUNTED, ACCOUNT_OPEN, ACCOUNT_OTHER, ACCOUNT_OUTSIDE, ACCOUNT_REMOVED, ACCOUNTS, MetricItems,
                                       Trace, metric_items, trace)

I1 = Iteration("I1", dt.datetime(2025, 5, 1), dt.datetime(2025, 5, 11))
I2 = Iteration("I2", dt.datetime(2025, 5, 11), dt.datetime(2025, 5, 21))
PIS = {"P1": (dt.datetime(2025, 5, 1), dt.datetime(2025, 6, 1)), "P2": (None, None)}          # a PI of the sheet without dates matches nothing
RULES = Rules.from_params()
AS_OF = dt.datetime(2025, 12, 1)


def day(text):
    return dt.datetime.fromisoformat(text)


def ticket(key, moves, team="ART"):
    changes, previous = [], "Nouveau"
    for n, (at, status) in enumerate(moves, start=1):
        changes.append(Change(f"{key}-{n}", day(at), previous, status))
        previous = status
    issue = FlowIssue(key=key, team=team, type="Feature", work_type="Feature", workflow="Feature", created=day("2025-04-01T00:00"),
                      status=previous, changes=tuple(changes))
    return item_flow(issue, AS_OF, RULES)


def traced(item):
    return trace(item, [I1, I2], PIS)


def test_the_accounts_are_five_and_named_once():
    assert ACCOUNTS == (ACCOUNT_COUNTED, ACCOUNT_OUTSIDE, ACCOUNT_OTHER, ACCOUNT_REMOVED, ACCOUNT_OPEN)
    assert len(set(ACCOUNTS)) == 5


def test_a_ticket_done_in_an_iteration_is_counted_with_its_iteration_and_pi():
    got = traced(ticket("A-1", [("2025-05-02T00:00", "En analyse"), ("2025-05-06T12:00", "Terminé")]))
    assert got == Trace("I1", day("2025-05-01"), day("2025-05-11"), "P1", ACCOUNT_COUNTED, "", 4.5 * 24)
    # nearest boundary: the start of I1 (05-01) is 5.5 days away, the end of I1 and the start of I2 (05-11) are 4.5 days away


def test_the_end_of_an_iteration_is_excluded_and_the_next_one_starts_there():
    got = traced(ticket("A-2", [("2025-05-02T00:00", "En analyse"), ("2025-05-11T00:00", "Terminé")]))
    assert (got.done_iteration, got.account, got.boundary_hours) == ("I2", ACCOUNT_COUNTED, 0.0)


def test_a_ticket_done_after_every_iteration_is_in_the_account_outside_and_has_no_iteration():
    got = traced(ticket("A-3", [("2025-05-02T00:00", "En analyse"), ("2025-05-25T00:00", "Terminé")]))
    assert (got.done_iteration, got.start, got.end, got.done_pi, got.account) == ("", None, None, "P1", ACCOUNT_OUTSIDE)    # in the PI, not in an iteration
    assert got.boundary_hours == 4 * 24                                                  # 05-25 is 4 days after the end of I2 (05-21)


def test_a_ticket_done_before_every_pi_has_no_pi():
    got = traced(ticket("A-4", [("2025-03-02T00:00", "En analyse"), ("2025-03-05T00:00", "Terminé")]))
    assert (got.done_iteration, got.done_pi, got.account) == ("", "", ACCOUNT_OUTSIDE)


def test_a_ticket_that_is_not_a_feature_is_left_out_even_when_done_in_an_iteration():
    got = traced(ticket("S-1", [("2025-05-02T00:00", "En analyse"), ("2025-05-06T00:00", "Terminé")], team=OTHER_TYPES))
    assert (got.done_iteration, got.account) == ("I1", ACCOUNT_OTHER)                    # left out of the ART metrics, but its iteration is known


def test_open_removed_and_never_started_tickets():
    assert traced(ticket("O-1", [("2025-05-02T00:00", "En analyse")])) == Trace("", None, None, "", ACCOUNT_OPEN, "", None)
    removed = traced(ticket("R-1", [("2025-05-02T00:00", "En analyse"), ("2025-05-03T00:00", "Annulé")]))
    assert (removed.account, removed.done_iteration, removed.boundary_hours) == (ACCOUNT_REMOVED, "", None)


def test_why_a_done_ticket_is_not_in_the_flow_time():
    straight = traced(ticket("W-1", [("2025-05-06T00:00", "Terminé")]))                  # Nouveau -> Terminé: no start
    reopened = traced(ticket("W-2", [("2025-05-02T00:00", "En analyse"), ("2025-05-03T00:00", "Terminé"),
                                     ("2025-05-04T00:00", "En développement"), ("2025-05-06T00:00", "Terminé")]))
    both = traced(ticket("W-3", [("2025-05-03T00:00", "Terminé"), ("2025-05-04T00:00", "En développement"), ("2025-05-06T00:00", "Terminé")]))
    clean = traced(ticket("W-4", [("2025-05-02T00:00", "En analyse"), ("2025-05-06T00:00", "Terminé")]))
    assert (straight.why_not_in_flow_time, reopened.why_not_in_flow_time, both.why_not_in_flow_time, clean.why_not_in_flow_time) == (
        "No start after the backlog", "Reopened", "Reopened, No start after the backlog", "")


def test_a_done_ticket_whose_workflow_has_no_start_status_says_so():
    """AL, 2026-11-22: the two causes of a missing start are told apart. No start status at all in the rules of the workflow (Jira's workflows
    not read, none declared in the sheet Workflows) is a setting to fix; a ticket that skipped the start statuses is a fact of the data."""
    no_start = Rules({k: StatusRule(r.category, False, r.is_end, r.is_free) for k, r in RULES._rows.items()})
    issue = FlowIssue(key="N-1", team="ART", type="Feature", work_type="Feature", workflow="Feature", created=day("2025-04-01T00:00"),
                      status="Terminé", changes=(Change("N-1-1", day("2025-05-02T00:00"), "Nouveau", "En analyse"),
                                                  Change("N-1-2", day("2025-05-06T00:00"), "En analyse", "Terminé")))
    item = item_flow(issue, AS_OF, no_start)
    assert (item.start_rule, item.skipped_start, traced(item).why_not_in_flow_time) == (False, True, "No start status in the rules")
    assert item_flow(issue, AS_OF, RULES).start_rule is True


def test_the_five_accounts_make_a_partition_of_any_population():
    population = [ticket("A-1", [("2025-05-02T00:00", "En analyse"), ("2025-05-06T00:00", "Terminé")]),
                  ticket("A-3", [("2025-05-02T00:00", "En analyse"), ("2025-05-25T00:00", "Terminé")]),
                  ticket("S-1", [("2025-05-02T00:00", "En analyse"), ("2025-05-06T00:00", "Terminé")], team=OTHER_TYPES),
                  ticket("R-1", [("2025-05-02T00:00", "En analyse"), ("2025-05-03T00:00", "Annulé")]),
                  ticket("O-1", [("2025-05-02T00:00", "En analyse")]), ticket("O-2", [])]
    accounts = [traced(i).account for i in population]
    assert accounts == [ACCOUNT_COUNTED, ACCOUNT_OUTSIDE, ACCOUNT_OTHER, ACCOUNT_REMOVED, ACCOUNT_OPEN, ACCOUNT_OPEN]
    assert all(a in ACCOUNTS for a in accounts)


# --- the lists behind the figures (sheet MetricItems) --------------------------------------------------------------------------------

def lists(*items):
    return metric_items(list(items), [I1, I2], PIS)


def population():
    return [ticket("A-1", [("2025-05-02T00:00", "En analyse"), ("2025-05-06T00:00", "Terminé")]),
            ticket("W-2", [("2025-05-02T00:00", "En analyse"), ("2025-05-03T00:00", "Terminé"), ("2025-05-04T00:00", "En développement"),
                           ("2025-05-06T00:00", "Terminé")]),
            ticket("C-1", [("2025-05-09T00:00", "En analyse"), ("2025-05-10T20:00", "Terminé")]),
            ticket("B-2", [("2025-05-12T00:00", "En analyse"), ("2025-05-13T00:00", "Terminé")]),
            ticket("A-3", [("2025-05-02T00:00", "En analyse"), ("2025-05-25T00:00", "Terminé")]),
            ticket("S-1", [("2025-05-02T00:00", "En analyse"), ("2025-05-06T00:00", "Terminé")], team=OTHER_TYPES),
            ticket("O-1", [("2025-05-02T00:00", "En analyse")])]


def test_the_list_behind_the_velocity_has_one_line_per_ticket_done_in_an_iteration_with_the_flow_time_verdict():
    got = lists(*population()).done
    s1, e1, s2, e2 = day("2025-05-01"), day("2025-05-11"), day("2025-05-11"), day("2025-05-21")
    # ActiveDays: the days in an Active status from StartUtc to DoneUtc (En analyse, En développement); W-2 was 1 day in Terminé between (not active)
    assert got == [("I1", s1, e1, "ART", "A-1", "Feature", day("2025-05-06T00:00"), 1, 4.0, "", 4.0, 0, 0, "Feature"),
                   ("I1", s1, e1, "ART", "W-2", "Feature", day("2025-05-06T00:00"), 0, 4.0, "Reopened", 3.0, 0, 0, "Feature"),
                   ("I1", s1, e1, "ART", "C-1", "Feature", day("2025-05-10T20:00"), 1, pytest.approx(1 + 20 / 24), "", pytest.approx(1 + 20 / 24), 0, 0,
                    "Feature"),
                   ("I2", s2, e2, "ART", "B-2", "Feature", day("2025-05-13T00:00"), 1, 1.0, "", 1.0, 0, 0, "Feature")]


def test_the_done_tickets_that_no_figure_counts_are_listed_with_their_reason():
    assert lists(*population()).left_out == [
        ("A-3", "Feature", "ART", "", day("2025-05-25T00:00"), ACCOUNT_OUTSIDE),
        ("S-1", "Feature", OTHER_TYPES, "I1", day("2025-05-06T00:00"), ACCOUNT_OTHER)]


def test_the_tickets_done_within_a_day_of_a_boundary_are_listed_by_distance_and_a_left_out_one_is_not():
    near = ticket("N-1", [("2025-05-02T00:00", "En analyse"), ("2025-05-10T22:30", "Terminé")], team=OTHER_TYPES)     # 1.5 h before the end of I1, left out
    assert lists(*population(), near).cutoff == [("C-1", "ART", "I1", day("2025-05-10T20:00"), 4.0)]


def test_an_empty_population_gives_three_empty_lists():
    assert lists() == MetricItems([], [], [])


def test_the_three_lists_become_the_rows_of_the_sheet_of_the_python_path_side_by_side_padded_with_empty_cells():
    from jira2dashboard.flow import formulas as F
    from jira2dashboard.standalone.results import metricitems_grid
    header, rows = metricitems_grid(lists(*population()))
    names = [[n for n, _ in block] for block in F.METRICITEMS_BLOCKS]
    assert header == tuple(names[0] + [""] + names[1] + [""] + names[2] + [""] + names[3])    # a blank column between the blocks
    assert len(rows) == 4 and all(len(r) == len(header) for r in rows)                          # the longest list (4 tickets done) sets the rows
    done_width, left_width = len(names[0]), len(names[1])
    assert [r[4] for r in rows] == ["A-1", "W-2", "C-1", "B-2"]                                 # the Key column of the first list
    assert [r[done_width + 1] for r in rows] == ["A-3", "S-1", None, None]                      # the second list, shorter: padded
    assert rows[0][done_width + 1 + left_width + 1] == "C-1" and rows[1][done_width + 1 + left_width + 1] is None   # the third list: one ticket


def test_the_position_of_a_column_of_a_list_follows_the_widths_of_the_blocks_before_it():
    from jira2dashboard.flow import formulas as F
    from jira2dashboard.flow import reconciliation as R
    widths = [len(b) for b in F.METRICITEMS_BLOCKS]
    assert widths == [14, 6, 5, 6]
    assert F.metricitems_column(3, "Iteration") == 29 and F.metricitems_cell(3, "Iteration") == "MetricItems_from_excel!$AC$3#"
    assert F.RECONCILIATION_FORMULAS[R.FLOW_LOAD_LIST][1] == '=IF(MetricItems_from_excel!$AC$3="",0,ROWS(MetricItems_from_excel!$AC$3#))'
    assert [text != "" for _name, text in F.METRICITEMS_AT_END] == [True] + [False] * 5     # one spill from the first column
    assert [F.metricitems_column(0, "Key"), F.metricitems_column(1, "Key"), F.metricitems_column(2, "Key")] == [5, 16, 23]
    assert F.metricitems_cell(2, "Key") == "MetricItems_from_excel!$W$3#" and F.metricitems_cell(0, "FlowTimeDays") == "MetricItems_from_excel!$I$3#"
    assert F.RECONCILIATION_FORMULAS[R.CUTOFF][1] == '=COUNTIF(MetricItems_from_excel!$W$3#,"?*")'        # the Key of the cut-off list


def test_the_first_list_is_sorted_by_the_start_of_the_iteration_before_the_team():
    late_a = ticket("Z-1", [("2025-05-12T00:00", "En analyse"), ("2025-05-13T00:00", "Terminé")], team="AAA")      # I2, a team that sorts first
    early_z = ticket("Z-2", [("2025-05-02T00:00", "En analyse"), ("2025-05-03T00:00", "Terminé")], team="ZZZ")     # I1, a team that sorts last
    assert [(r[0], r[3], r[4]) for r in lists(late_a, early_z).done] == [("I1", "ZZZ", "Z-2"), ("I2", "AAA", "Z-1")]
