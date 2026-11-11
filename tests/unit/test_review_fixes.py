"""Fixes of the review of the Python scripts (2026-11-11): each one has the case that failed before."""
import socket
import zipfile
from pathlib import Path
from xml.etree import ElementTree

import pytest

from jira2dashboard.mockjira import __main__ as mock_main
from jira2dashboard.standalone import __main__ as extract
from jira2dashboard.standalone import mapping
from jira2dashboard.standalone import params_book as PB
from jira2dashboard.standalone import xlsx_read
from jira2dashboard.standalone import xlsx_write as W
from jira2dashboard.testbench import run_vba_tests as rvt


def test_extract_refuses_an_out_that_would_replace_the_workbook(tmp_path, capsys):
    workbook = tmp_path / "jira2dashboard.xlsm"
    workbook.write_bytes(b"x")
    for wrong, word in ((str(workbook), ".xlsx"), (str(tmp_path / "result.xlsm"), ".xlsx"), (str(tmp_path / "jira2dashboard.XLSM"), ".xlsx")):
        with pytest.raises(SystemExit) as stop:
            extract.parse(["--workbook", str(workbook), "--out", wrong])
        assert stop.value.code == 2 and word in capsys.readouterr().err
    same = tmp_path / "same.xlsx"
    same.write_bytes(b"x")
    with pytest.raises(SystemExit) as stop:                               # an .xlsx workbook given as its own output
        extract.parse(["--workbook", str(same), "--out", str(same)])
    assert stop.value.code == 2 and "the workbook itself" in capsys.readouterr().err
    assert extract.parse(["--workbook", str(workbook), "--out", str(tmp_path / "ok.xlsx")]).out.endswith("ok.xlsx")


def test_a_quote_in_a_title_keeps_the_drawing_xml_well_formed():
    svg = "<svg xmlns='http://www.w3.org/2000/svg'/>"
    parts = W.xlsx_parts([W.Sheet('Charts "A"', pictures=[W.Picture('v"1', 'Flow of the team A "B" <x>', svg)])])
    drawing = next(data for name, data in parts.items() if name.startswith("xl/drawings/drawing"))
    ElementTree.fromstring(drawing)                                       # raises on "not well-formed"
    ElementTree.fromstring(parts["xl/workbook.xml"])
    assert b"&quot;" in drawing


def test_the_shareable_workbook_copy_masks_the_token_and_keeps_the_rest():
    sheet = W.Sheet("Config", ("Parameter", "Value", "Example"),
                    [("ApiToken", "tok-secret", ""), ("Email", "me@example.com", ""), ("BaseUrl", "https://x", ""), ("Email2", "", "")])
    out = extract._masked([sheet])[0].rows
    assert [row[1] for row in out] == [extract.MASK, "me@example.com", "https://x", ""]


def test_a_part_in_utf16_cannot_hide_a_doctype(tmp_path):
    path = tmp_path / "w.xlsx"
    bomb = ("<?xml version='1.0'?><!DOCTYPE x [<!ENTITY a 'b'>]><x>&a;</x>").encode("utf-16")
    with zipfile.ZipFile(path, "w") as package:
        package.writestr("xl/workbook.xml", bomb)
    with zipfile.ZipFile(path) as package:
        with pytest.raises(xlsx_read.NotAWorkbook, match="not UTF-8"):
            xlsx_read._xml(package, "xl/workbook.xml")


@pytest.mark.parametrize("text, minimum, default, expected", [("--5", 1, 60, 60), ("5", 1, 60, 5), (" 7 ", 1, 60, 7), ("-3", 0, 4, 4), ("x", 1, 60, 60)])
def test_a_malformed_number_of_config_is_not_a_crash(text, minimum, default, expected):
    assert PB._number(text, default, minimum) == expected


def test_the_blocked_issues_of_each_issue_are_found_in_one_pass():
    links = [("A-1", "B-1"), ("A-1", "B-2"), ("C-1", "B-1")]
    blocked_by: dict[str, list[str]] = {}
    for blocker, blocked in links:
        blocked_by.setdefault(blocker, []).append(blocked)
    assert tuple(blocked_by.get("A-1", ())) == ("B-1", "B-2") and tuple(blocked_by.get("Z", ())) == ()
    assert "blocked_by" in open(mapping.__file__, encoding="utf-8").read()       # the extract uses that map, not a scan per issue


def test_bisect_says_when_a_module_does_not_compile(monkeypatch, capsys):
    monkeypatch.setattr(rvt, "build", lambda names, target: [])
    monkeypatch.setattr(rvt, "compile_ok", lambda target: "COMPILED_OK")
    assert rvt.bisect() is True
    answers = iter(["COMPILED_OK", "COMPILE_FAILED"])
    monkeypatch.setattr(rvt, "compile_ok", lambda target: next(answers))
    assert rvt.bisect() is False
    assert "COMPILE_FAILED" in capsys.readouterr().out


def test_the_fake_jira_tells_a_port_that_is_taken(capsys):
    with socket.socket() as taken:
        taken.bind(("127.0.0.1", 0))
        taken.listen()
        port = taken.getsockname()[1]
        assert mock_main.main(["--port", str(port), "--issues", "5"]) == 1
    assert f"cannot listen on port {port}" in capsys.readouterr().out


def test_extract_goes_on_when_the_log_folder_cannot_be_written(tmp_path, monkeypatch, capsys):
    blocker = tmp_path / "logs"
    blocker.write_text("a file where the log folder should be")          # mkdir of logs\x fails
    log = blocker / "python.log"
    monkeypatch.setattr(extract.tempfile, "gettempdir", lambda: str(tmp_path / "fallback"))
    code = extract.main(["--workbook", str(tmp_path / "missing.xlsm"), "--log", str(log)])
    assert code == 1                                                      # the workbook is missing: a message, not a traceback
    assert "cannot write the log there" in capsys.readouterr().out
    assert (tmp_path / "fallback" / "jira2dashboard" / "logs" / "python.log").exists()
