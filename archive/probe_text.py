"""Probe: which texts does French Excel convert when Python writes them with Range.Value, with and without "'"."""
import xlsm_ref as R

TEXTS = ["true", "TRUE", "1.10", "1,10", "2026-10-01", "1/2", "007", "50%", "#N/A", "=1+1", "-x", "PI3-S1", "abc"]

with R.excel_session() as app:
    wb = R.new_workbook(app)
    ws = wb.Worksheets(1)
    R.write_block(ws, 1, 1, [TEXTS])
    R.write_block(ws, 2, 1, [["'" + t for t in TEXTS]])
    print(f"{'written':12} {'plain -> value (type)':32} {'with apostrophe -> value (type)'}")
    for c, t in enumerate(TEXTS, start=1):
        a, b = ws.Cells(1, c).Value, ws.Cells(2, c).Value
        print(f"{t!r:12} {repr(a) + ' (' + type(a).__name__ + ')':32} {b!r} ({type(b).__name__})")
    wb.Close(False)
    del ws, wb
