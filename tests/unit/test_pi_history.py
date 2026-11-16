"""The PI field through time (flow/pi_history.py): hand-written cases."""
import datetime as dt

from jira2dashboard.flow.model import FlowIssue, PIChange
from jira2dashboard.flow.pi_history import PIInterval, pi_at, pi_intervals

CREATED = dt.datetime(2026, 1, 1)


def at(day: int, hour: int = 0) -> dt.datetime:
    return dt.datetime(2026, 1, day, hour)


def issue(pi, *changes):
    return FlowIssue(key="F-1", team="ART", type="Feature", work_type="Feature", workflow="Feature", created=CREATED,
                     status="Nouveau", pi=pi, pi_changes=tuple(PIChange(*c) for c in changes))


def test_without_any_change_the_current_value_holds_from_the_creation():
    found = pi_intervals(issue("PI-1"))
    assert found == [PIInterval("F-1", "PI-1", CREATED, None)]
    assert pi_at(found, CREATED) == "PI-1" and pi_at(found, dt.datetime(2030, 1, 1)) == "PI-1"
    assert pi_at(found, CREATED - dt.timedelta(seconds=1)) is None            # before the issue existed


def test_the_value_before_the_first_change_is_the_from_of_that_change_not_the_current_value():
    found = pi_intervals(issue("PI-2", ("1", at(5), None, "PI-2")))
    assert found == [PIInterval("F-1", "PI-2", at(5), None)]                  # empty from the creation to the change
    assert pi_at(found, at(4)) is None and pi_at(found, at(6)) == "PI-2"


def test_a_pi_emptied_or_moved_ends_the_stretch_and_the_end_is_excluded_the_start_included():
    emptied = pi_intervals(issue(None, ("1", at(3), "PI-1", None)))
    assert emptied == [PIInterval("F-1", "PI-1", CREATED, at(3))]
    assert pi_at(emptied, at(3) - dt.timedelta(seconds=1)) == "PI-1" and pi_at(emptied, at(3)) is None
    moved = pi_intervals(issue("PI-5", ("1", at(3), "PI-3", "PI-5")))
    assert [(i.pi, i.start, i.end) for i in moved] == [("PI-3", CREATED, at(3)), ("PI-5", at(3), None)]
    assert pi_at(moved, at(3)) == "PI-5"


def test_two_changes_in_the_same_second_are_ordered_by_their_id_as_a_number():
    changes = (("10", at(3), "PI-2", "PI-3"), ("9", at(3), "PI-1", "PI-2"))      # 9 before 10, as Jira gave them
    found = pi_intervals(issue("PI-3", *changes))
    assert [(i.pi, i.start, i.end) for i in found] == [("PI-1", CREATED, at(3)), ("PI-3", at(3), None)]


def test_blanks_around_a_pi_are_dropped_and_a_blank_value_is_no_pi():
    found = pi_intervals(issue("  PI-1 ", ("1", at(2), "  ", " PI-1 ")))
    assert [(i.pi, i.start) for i in found] == [("PI-1", at(2))]
    assert pi_intervals(issue("")) == [] and pi_intervals(issue(None)) == []


def test_a_change_dated_before_the_creation_does_not_make_an_interval_before_it():
    found = pi_intervals(issue("PI-2", ("1", CREATED - dt.timedelta(days=1), "PI-1", "PI-2")))
    assert found == [PIInterval("F-1", "PI-2", CREATED, None)]
