"""The code of the workbook in one text file for Rovo (src/jira2dashboard/vba_code.py, AL 2026-11-27): the rows of the hidden sheet VbaCode and the text
that the button VBA code for Rovo writes. The VBA writer is checked against `file_text` by vba_code.feature (the real workbook) and T_VbaCode."""
from jira2dashboard import build_workbook
from jira2dashboard import vba_code as VC
from jira2dashboard.paths import VBA_DIR


def test_the_lines_of_a_module_lose_only_the_name_attribute_and_the_last_newline():
    text = 'Attribute VB_Name = "modA"\nOption Explicit\n\n    \' a comment\n'
    assert VC.module_lines(text) == ["Option Explicit", "", "    ' a comment"]
    assert VC.module_lines("a\r\nb") == ["a", "b"]


def test_the_rows_number_the_lines_of_each_module_from_1_in_the_order_given():
    assert VC.rows([("modB", "x\ny\n"), ("modA", "z\n")]) == [("modB", 1, "x"), ("modB", 2, "y"), ("modA", 1, "z")]


def test_the_text_holds_the_head_each_module_under_its_title_then_the_formulas():
    rows = [("modA", 1, "Option Explicit"), ("modA", 2, None), ("modB", 1, "Sub X()")]
    text = VC.file_text(rows, [("ItemFlow", "Key", ""), ("ItemFlow", "IsDone", "=1"), ("T", "C", None)], "v1 commit abc")
    assert text == "\n".join([
        VC.TITLE, "Version: v1 commit abc", VC.INTRO, "Modules: modA (2 lines), modB (1 lines)",
        "", "===== VBA module modA (2 lines) =====", "Option Explicit", "",
        "", "===== VBA module modB (1 lines) =====", "Sub X()",
        "", VC.FORMULAS_HEAD, "ItemFlow | Key | ", "ItemFlow | IsDone | =1", "T | C | "]) + "\n"


def test_the_delivered_modules_are_those_of_the_build_without_the_test_modules():
    assert build_workbook.MODULES is VC.MODULES
    assert not {"modTests", "modButtons", "modBackground"} & set(VC.MODULES)
    rows = VC.rows(VC.delivered(VBA_DIR, VC.MODULES))
    assert {r[0] for r in rows} == set(VC.MODULES) and all(len(r[2]) <= 32766 for r in rows)            # a cell holds 32 767 characters
    assert any("Public Sub BtnExtract()" in r[2] for r in rows) and not any(r[2].startswith("Attribute VB_Name") for r in rows)
