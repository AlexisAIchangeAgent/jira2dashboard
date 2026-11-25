"""The order of the PIs (AL, 2026-11-25): by Start; the PIs without dates in the natural order of their names (the numbers as numbers), and any
name is a PI. The same names and the same order, typed by hand, are in modTests.T_PiSortKey: the VBA key sorts them the same way."""
import datetime as dt

from jira2dashboard.flow.model import ordered_pis, pi_sort_key

NAMES = ["PI#10-2026", "pi#05-2026", "PI#9-2027", "Champ_customefield_11339", "PI#4.2-2024", "PI#1-2027", "PI 4", "PI#3-2024", "PI#9-2026",
         "PI#01-2027", "PI#4.1-2024"]
NATURAL = ["Champ_customefield_11339", "PI 4", "PI#01-2027", "PI#1-2027", "PI#3-2024", "PI#4.1-2024", "PI#4.2-2024", "pi#05-2026", "PI#9-2026",
           "PI#9-2027", "PI#10-2026"]


def test_the_names_without_dates_come_in_the_natural_order():
    assert sorted(NAMES, key=lambda name: (pi_sort_key(name), name)) == NATURAL


def test_the_key_is_the_text_of_the_vba():
    """A run of digits = character 1, its length on 4 digits, the digits without their leading zeros; a to z in capitals, the rest as it is."""
    assert pi_sort_key("pi#05-2026") == "PI#\x0100015-\x0100042026"
    assert pi_sort_key("PI#01-2027") == pi_sort_key("PI#1-2027") and pi_sort_key("a00") == "A\x010000" and pi_sort_key("") == ""


def test_the_pis_with_dates_come_by_start_then_by_name():
    day = dt.datetime
    pis = {"PI#10-2026": (day(2026, 1, 1), day(2026, 2, 1)), "PI#9-2026": (day(2026, 1, 1), day(2026, 3, 1)),
           "PI#1-2025": (day(2026, 2, 1), day(2026, 4, 1)), "Early": (day(2025, 1, 1), day(2025, 2, 1))}
    assert [name for name, _window in ordered_pis(pis)] == ["Early", "PI#9-2026", "PI#10-2026", "PI#1-2025"]
