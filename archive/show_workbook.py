"""Describe the generated test workbook: sheets, VBA modules, first result rows (macros disabled)."""
from pathlib import Path
import xlsm_ref as R

path = Path(__file__).parent / "out" / "vba_tests.xlsm"
print("File:", path, f"({path.stat().st_size // 1024} KB)")
with R.excel_session() as app:
    wb = R.open_without_macros(app, path)
    print("FileFormat:", int(wb.FileFormat), "(52 = .xlsm)")
    print("\nSheets:")
    for i in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(i)
        used = ws.UsedRange
        print(f"  {ws.Name:14} {used.Rows.Count:4} rows x {used.Columns.Count} cols")
    print("\nVBA modules:")
    comps = wb.VBProject.VBComponents
    for i in range(1, comps.Count + 1):
        c = comps.Item(i)
        kind = {1: "module", 2: "class", 3: "form", 100: "document"}.get(c.Type, c.Type)
        print(f"  {c.Name:14} {kind:9} {c.CodeModule.CountOfLines:4} lines")
    rows = wb.Worksheets("TestResults").UsedRange.Value
    print("\nTestResults (first 8 rows of", len(rows) - 1, "):")
    for r in rows[:9]:
        print("  " + " | ".join(str(x)[:28] for x in r[:7]))
    print("\nCoverage (CoverageMap!E2):", wb.Worksheets("CoverageMap").Range("E2").Value2)
    wb.Close(False)
    del ws, used, comps, c, wb
