"""Characterization tests: pure helpers of excel/com.py (no Excel needed)."""
import datetime as dt
from pathlib import Path

import pytest
import pywintypes

from jira2dashboard.excel import com


# --- as_cell: text gets a leading apostrophe so Excel keeps it as text ---------------------------------

@pytest.mark.parametrize("text", ["true", "1.10", "2026-10-01", "007", "50%", "=1+1", "abc"])
def test_as_cell_prefixes_text_with_an_apostrophe(text):
    assert com.as_cell(text) == "'" + text


@pytest.mark.parametrize("value", [None, 0, 3.5, True, dt.date(2026, 11, 3), ""])
def test_as_cell_leaves_non_text_and_empty_text_alone(value):
    assert com.as_cell(value) is value or com.as_cell(value) == value


def test_as_cell_truncates_to_the_excel_cell_limit():
    result = com.as_cell("x" * 40000)
    assert result == "'" + "x" * 32766


# --- to_excel_serial --------------------------------------------------------------------------------

def test_to_excel_serial_of_a_date():
    assert com.to_excel_serial(dt.date(2026, 11, 3)) == 46329


def test_to_excel_serial_keeps_the_time_of_day():
    assert com.to_excel_serial(dt.datetime(2026, 11, 3, 12, 0)) == 46329.5


def test_to_excel_serial_epoch_quirk_1900():
    assert com.to_excel_serial(dt.date(1900, 3, 1)) == 61


def test_to_excel_serial_refuses_time_zone_aware_values():
    with pytest.raises(ValueError):
        com.to_excel_serial(dt.datetime(2026, 11, 3, tzinfo=dt.timezone.utc))


# --- to_2d_tuple ------------------------------------------------------------------------------------

def test_to_2d_tuple_returns_a_tuple_of_tuples():
    assert com.to_2d_tuple([[1, 2], [3, 4]]) == ((1, 2), (3, 4))


def test_to_2d_tuple_accepts_no_rows():
    assert com.to_2d_tuple([]) == ()


def test_to_2d_tuple_refuses_rows_of_different_lengths():
    with pytest.raises(ValueError, match="different lengths"):
        com.to_2d_tuple([[1, 2], [3]])


# --- VBA text helpers -------------------------------------------------------------------------------

def test_normalize_vba_strips_attribute_lines_and_uses_crlf():
    code = 'Attribute VB_Name = "modX"\nOption Explicit\n\nPublic Sub A()\nEnd Sub\n\n\n'
    assert com.normalize_vba(code) == "Option Explicit\r\n\r\nPublic Sub A()\r\nEnd Sub\r\n"


def test_normalize_vba_handles_mixed_line_endings_and_trailing_spaces():
    assert com.normalize_vba("a  \r\nb\rc\n") == "a\r\nb\r\nc\r\n"


def test_normalize_vba_drops_leading_blank_lines():
    assert com.normalize_vba("\n\nSub A()\nEnd Sub") == "Sub A()\r\nEnd Sub\r\n"


def test_normalize_vba_is_idempotent():
    once = com.normalize_vba("Attribute VB_Name = \"m\"\nSub A()\nEnd Sub")
    assert com.normalize_vba(once) == once


@pytest.mark.parametrize("name", ["modMain", "a", "A_1", "m" * 31])
def test_valid_module_names(name):
    assert com.is_valid_module_name(name)


@pytest.mark.parametrize("name", ["", "1x", "_x", "mod Main", "m" * 32, "modCafé"])
def test_invalid_module_names(name):
    assert not com.is_valid_module_name(name)


def test_vba_string_literal_doubles_quotes_and_splits_lines():
    assert com.vba_string_literal('a"b\nc') == '"a""b" & vbLf & "c"'


def test_vba_string_literal_plain_text():
    assert com.vba_string_literal("abc") == '"abc"'


# --- COM error helpers ------------------------------------------------------------------------------

def _com_error(hresult, scode=None):
    excepinfo = (0, "Excel", "text", None, 0, scode) if scode is not None else None
    return pywintypes.com_error(hresult, "msg", excepinfo, None)


def test_com_scode_reads_the_inner_scode():
    assert com.com_scode(_com_error(com.DISP_E_EXCEPTION, com.XL_ERROR_1004)) == com.XL_ERROR_1004


def test_com_scode_is_none_without_excepinfo():
    assert com.com_scode(_com_error(com.DISP_E_EXCEPTION)) is None


def test_call_with_retry_retries_while_excel_is_busy():
    calls = []

    def flaky():
        calls.append(1)
        if len(calls) < 3:
            raise _com_error(com.RPC_E_CALL_REJECTED)
        return "ok"

    assert com.call_with_retry(flaky, delay_s=0) == "ok"
    assert len(calls) == 3


def test_call_with_retry_does_not_retry_other_errors():
    calls = []

    def broken():
        calls.append(1)
        raise _com_error(com.DISP_E_EXCEPTION)

    with pytest.raises(pywintypes.com_error):
        com.call_with_retry(broken, delay_s=0)
    assert len(calls) == 1


def test_call_with_retry_gives_up_after_the_last_attempt():
    calls = []

    def busy():
        calls.append(1)
        raise _com_error(com.RPC_E_SERVERCALL_RETRYLATER)

    with pytest.raises(pywintypes.com_error):
        com.call_with_retry(busy, attempts=3, delay_s=0)
    assert len(calls) == 3


# --- VBA trust (registry reads are replaced) --------------------------------------------------------

def test_policy_value_wins_over_user_value(monkeypatch):
    values = {com.POLICY_KEY: 0, com.USER_KEY: 1}
    monkeypatch.setattr(com, "read_access_vbom", values.get)
    assert com.is_vbom_trusted() is False


def test_user_value_is_used_without_policy(monkeypatch):
    values = {com.POLICY_KEY: None, com.USER_KEY: 1}
    monkeypatch.setattr(com, "read_access_vbom", values.get)
    assert com.is_vbom_trusted() is True


def test_not_trusted_when_nothing_is_set(monkeypatch):
    monkeypatch.setattr(com, "read_access_vbom", lambda key: None)
    assert com.is_vbom_trusted() is False


# --- save_as_xlsm with a fake workbook --------------------------------------------------------------

class FakeWorkbook:
    def __init__(self):
        self.saved = None

    def SaveAs(self, path, file_format):
        self.saved = (path, file_format)
        Path(path).write_bytes(b"x")


def test_save_as_xlsm_refuses_another_extension(tmp_path):
    with pytest.raises(ValueError, match=".xlsm"):
        com.save_as_xlsm(FakeWorkbook(), tmp_path / "a.xlsx")


def test_save_as_xlsm_refuses_to_overwrite_by_default(tmp_path):
    target = tmp_path / "a.xlsm"
    target.write_bytes(b"old")
    with pytest.raises(FileExistsError):
        com.save_as_xlsm(FakeWorkbook(), target)
    assert target.read_bytes() == b"old"


def test_save_as_xlsm_overwrites_when_asked_and_uses_format_52(tmp_path):
    target = tmp_path / "sub" / "a.xlsm"
    target.parent.mkdir()
    target.write_bytes(b"old")
    wb = FakeWorkbook()
    result = com.save_as_xlsm(wb, target, overwrite=True)
    assert result == target.resolve()
    assert wb.saved == (str(target.resolve()), 52)


def test_save_as_xlsm_creates_missing_folders(tmp_path):
    target = tmp_path / "new" / "deeper" / "a.xlsm"
    com.save_as_xlsm(FakeWorkbook(), target)
    assert target.exists()
