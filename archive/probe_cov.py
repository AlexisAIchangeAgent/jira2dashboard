"""Probe the CoverageMap formulas in the saved test workbook."""
from pathlib import Path
import xlsm_ref as R

with R.excel_session() as app:
    wb = R.open_without_macros(app, Path(__file__).parent / "out" / "vba_tests.xlsm")
    ws = wb.Worksheets("CoverageMap")
    print("Calculation mode:", app.Calculation, "(-4105 = automatic, -4135 = manual)")
    print("E2 Value2:", ws.Range("E2").Value2, "Text:", ws.Range("E2").Text, "Formula:", ws.Range("E2").Formula)
    for r in (2, 3, 30, 66):
        print(r, ws.Cells(r, 1).Value, ws.Cells(r, 2).Value2, ws.Cells(r, 3).Value2, ws.Cells(r, 2).Formula)
    hits = wb.Worksheets("CoverageHits")
    print("CoverageHits rows:", hits.UsedRange.Rows.Count, "A2:", repr(hits.Range("A2").Value), "B2:", hits.Range("B2").Value)
    app.CalculateFull()
    print("after CalculateFull -> E2:", ws.Range("E2").Value2)
    wb.Close(False)
    del ws, hits, wb
