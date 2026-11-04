"""Check the progress bar and the row order in the saved v3 workbook (macros disabled)."""
from pathlib import Path
import xlsm_ref as R

with R.excel_session() as app:
    wb = R.open_without_macros(app, Path(__file__).parent / "out" / "vba_tests_v3.xlsm")
    ws = wb.Worksheets("TestResults")
    back, fill = ws.Shapes("pbBack"), ws.Shapes("pbFill")
    print(f"pbBack width={back.Width:.1f}  pbFill width={fill.Width:.1f}  label J10={ws.Range('J10').Value!r}")
    times = [ws.Cells(r, 8).Value2 for r in range(2, 132)]
    print("RunAt first/last row (seconds apart):", round((times[-1] - times[0]) * 86400, 1))
    print("rows:", ws.Cells(65536, 1).End(-4162).Row - 1, "| last row:", ws.Range("A131").Value, ws.Range("G131").Value)
    wb.Close(False)
    del back, fill, ws, wb
