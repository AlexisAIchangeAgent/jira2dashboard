"""The support codes: the table of error_codes.py is the reference, modLog.ErrorCode (VBA) and the sheet ErrorCodes follow it."""
import re
from pathlib import Path

import pytest

from jira2dashboard import error_codes as E

VBA = Path(__file__).resolve().parents[2] / "vba"


def vba_cases() -> dict[int, str]:
    """`Case 608: ErrorCode = "J2D-CFG-002"` of modLog.ErrorCode, by the number raised as vbObjectError + n."""
    text = (VBA / "modLog.bas").read_text(encoding="ascii")
    return {int(n): code for n, code in re.findall(r'Case (\d+): ErrorCode = "(J2D-[A-Z]+-\d+)"', text)}


def raised_numbers() -> set[int]:
    found = set()
    for path in VBA.glob("*.bas"):
        if path.name == "modTests.bas":
            continue
        found |= {int(n) for n in re.findall(r"Err\.Raise vbObjectError \+ (\d+)", path.read_text(encoding="ascii"))}
    return found


def test_every_code_has_the_format_and_is_unique_and_filled():
    codes = [row.code for row in E.TABLE]
    assert len(codes) == len(set(codes))
    assert all(E.CODE_FORMAT.match(code) for code in codes), [c for c in codes if not E.CODE_FORMAT.match(c)]
    assert all(all(cell.strip() for cell in row[2:]) and row.meaning for row in E.TABLE)       # no empty cause, check or support text


def test_the_vba_table_is_the_python_table():
    expected = {row.vba: row.code for row in E.TABLE if row.vba is not None and row.vba != 602}
    assert vba_cases() == expected
    assert E.BY_CODE["J2D-CFG-005"].vba == 602                      # 602 has two causes: the BaseUrl check, and an HTTP status


def test_every_error_the_vba_raises_has_a_code():
    missing = raised_numbers() - set(vba_cases()) - {602}
    assert not missing, f"vbObjectError + {sorted(missing)} raised without a code in modLog.ErrorCode and in error_codes.TABLE"


@pytest.mark.parametrize("status, code", [(401, "J2D-JIRA-401"), (503, "J2D-JIRA-503"), (404, "J2D-JIRA-404"), (0, "J2D-NET-001")])
def test_an_http_status_gives_its_code(status, code):
    assert E.code_of_http(status) == code


@pytest.mark.parametrize("number, code", [(0x80072EE7 - 2**32, "J2D-NET-12007"), (12029, "J2D-NET-12029"), (13, "J2D-NET-001"), (-2147467259, "J2D-NET-001")])
def test_a_winhttp_error_gives_its_code(number, code):
    assert E.code_of_winhttp(number) == code


def test_the_http_statuses_of_the_table_are_real_statuses():
    statuses = [int(row.code[-3:]) for row in E.TABLE if row.code.startswith("J2D-JIRA-") and int(row.code[-3:]) >= 400]
    assert statuses == sorted(statuses) and statuses[0] == 400


def test_the_sheet_rows_and_the_markdown_hold_every_code():
    assert len(E.rows()) == len(E.TABLE) and all(len(r) == len(E.HEADER) for r in E.rows())
    text = E.markdown()
    assert all(row.code in text for row in E.TABLE) and text.count("\n") == len(E.TABLE) + 1


# --- the Python side: the same code for the same cause ---------------------------------------------------------------------

def test_the_python_rule_gives_the_codes_of_the_vba_tests():
    """The cases of modTests.T_ErrorCodes, on the Python rule."""
    assert E.code_of_vba(608) == "J2D-CFG-002" and E.code_of_vba(610) == "J2D-SYS-001"
    assert E.code_of_vba(602, "HTTP 401 after 1 attempt(s).  {x}") == "J2D-JIRA-401"
    assert E.code_of_vba(602, "BaseUrl must start with https://") == "J2D-CFG-005"
    assert E.code_of_vba(602, "HTTP 0 after 3 attempt(s). -2147012889 The server name could not be resolved") == "J2D-NET-12007"
    assert E.code_of_vba(602, "HTTP 0 after 3 attempt(s). cancelled") == "J2D-NET-001"
    assert E.code_of_vba(999) == "J2D-GEN-999"


def test_a_jira_error_knows_its_support_code():
    from jira2dashboard.standalone import jira_api as J
    assert J.status_error(401, 1, "", "").support_code == "J2D-JIRA-401"
    assert J.status_error(0, 3, "-2147012889 The server name or address could not be resolved", "").support_code == "J2D-NET-12007"
    assert J.JiraError(603, "x").support_code == "J2D-CFG-003"
    assert J.JiraError(607, "x").support_code == "J2D-JIRA-020"


def test_the_log_line_of_an_error_starts_with_its_code(tmp_path):
    from jira2dashboard.standalone import __main__ as extract
    workbook = tmp_path / "missing.xlsm"
    log = tmp_path / "logs" / "python.log"
    assert extract.main(["--workbook", str(workbook), "--log", str(log)]) == 1
    text = log.read_text(encoding="utf-8")
    assert "ERROR J2D-SYS-011 | " in text and "RUN end stopped J2D-SYS-011 (log: " in text


def test_every_expected_error_of_the_extraction_has_a_code_of_the_table():
    from jira2dashboard.params.model import ParamsError
    from jira2dashboard.standalone import __main__ as extract
    from jira2dashboard.standalone.xlsx_write import XlsxLocked
    for err in (ParamsError("x"), ParamsError("x", "J2D-CFG-002"), XlsxLocked("x"), FileNotFoundError("x"), ValueError("x")):
        assert extract.support_code(err) in E.BY_CODE, err
    assert extract.UNEXPECTED_CODE in E.BY_CODE
