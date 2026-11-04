"""Integration: excel/com.py against a real Excel (marker `excel`, not in the fast loop)."""
import datetime as dt

import pytest

from jira2dashboard.excel import com

pytestmark = pytest.mark.excel


def test_write_block_round_trips_numbers_text_and_dates():
    with com.excel_session() as app:
        wb = com.new_workbook(app)
        ws = wb.Worksheets(1)
        day = dt.date(2026, 11, 3)
        com.write_block(ws, 1, 1, [["007", 2, com.to_excel_serial(day)], ["true", 3.5, None]])
        back = com.block_range(ws, 1, 1, 2, 3).Value
        assert back[0][0] == "007" and back[1][0] == "true"      # text kept as text
        assert back[0][1] == 2 and back[1][1] == 3.5 and back[1][2] is None
        assert ws.Range("C1").Value2 == com.to_excel_serial(day)
        wb.Close(False)
        del ws, wb, back, app


def test_english_formula_and_number_format_work_in_any_excel_language():
    with com.excel_session() as app:
        wb = com.new_workbook(app)
        ws = wb.Worksheets(1)
        ws.Range("A1").Value = 2
        ws.Range("A2").Value = 3.5
        cell = ws.Range("A3")
        cell.Formula = "=SUM(A1:A2)"
        com.put_en_us(cell, "NumberFormat", "0.00")
        assert cell.Value == 5.5
        assert cell.NumberFormatLocal in ("0,00", "0.00")
        wb.Close(False)
        del cell, ws, wb, app
