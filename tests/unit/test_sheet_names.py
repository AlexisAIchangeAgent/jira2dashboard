"""The names of the worksheets say where their data comes from: _from_user (typed by the person), _from_Jira (read from
the Jira API), _from_excel (calculated, plotted or written by the workbook itself). The code keeps the logical names
(`Config`, `Issues`...) and turns them into the real name where it meets Excel; the VBA does the same (modText.SheetName)."""
import re
from pathlib import Path

import pytest

from jira2dashboard import sheet_names as S

VBA = Path(__file__).resolve().parents[2] / "vba"


def test_every_sheet_name_ends_with_the_suffix_of_its_source():
    assert S.of("Config") == "Config_from_user"
    assert S.of("Issues") == "Issues_from_Jira"
    assert S.of("Metrics") == "Metrics_from_excel"
    assert S.of("JiraFields") == "JiraFields_from_Jira"      # a proposal read from Jira keeps its Jira prefix


def test_a_real_name_is_left_alone_and_gives_its_logical_name_back():
    assert S.of("Config_from_user") == "Config_from_user"
    assert S.logical("Config_from_user") == "Config"
    assert S.logical("Config") == "Config"


def test_a_name_the_workbook_does_not_know_is_left_alone():
    assert S.of("TestScratch") == "TestScratch"
    assert S.logical("TestScratch") == "TestScratch"


def test_no_sheet_is_in_two_sources_and_every_name_fits_the_31_characters_of_excel():
    every = S.USER + S.JIRA + S.EXCEL
    assert len(every) == len(set(every))
    assert max(len(S.of(name)) for name in every) <= 31
    assert all(S.of(name).endswith(S.SUFFIX[S.source(name)]) for name in every)


@pytest.mark.parametrize("name, source", [("Projects", "user"), ("PIs", "user"), ("StatusIntervals", "jira"),
                                          ("IssueLinks", "jira"), ("ItemFlow", "excel"), ("Dashboard", "excel"),
                                          ("Charts", "excel"), ("Manual", "excel"), ("RovoPromptSpec", "excel")])
def test_the_source_of_a_sheet(name, source):
    assert S.source(name) == source


def test_the_vba_lists_are_the_python_lists():
    # modText.SheetName: the same three lists, in the same order (a new sheet is added to both, or this fails)
    text = (VBA / "modText.bas").read_text(encoding="ascii")
    for const, names in (("USER_SHEETS", S.USER), ("JIRA_SHEETS", S.JIRA), ("EXCEL_SHEETS", S.EXCEL)):
        found = re.search(rf'Private Const {const} As String = "([^"]*)"', text)
        assert found, const
        assert found.group(1).strip("|").split("|") == list(names), const
    for const, suffix in (("SUFFIX_USER", S.FROM_USER), ("SUFFIX_JIRA", S.FROM_JIRA), ("SUFFIX_EXCEL", S.FROM_EXCEL)):
        assert re.search(rf'Private Const {const} As String = "{suffix}"', text), const
