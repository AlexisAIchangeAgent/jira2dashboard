"""Print selected rows of the saved TestResults sheet (macros disabled)."""
import sys
from pathlib import Path
import xlsm_ref as R

with R.excel_session() as app:
    wb = R.open_without_macros(app, Path(__file__).parent / "out" / "vba_tests.xlsm")
    rows = wb.Worksheets("TestResults").UsedRange.Value
    cov = wb.Worksheets("CoverageMap").Range("E2").Text
    wb.Close(False)
    del wb
areas = {}
for r in rows[1:]:
    areas.setdefault(r[1], [0, 0])[0 if r[6] == "OK" else 1] += 1
print("rows:", len(rows) - 1, "| CoverageMap!E2 =", cov)
for area, (ok, ko) in sorted(areas.items()):
    print(f"  {area:11} OK={ok:3}  KO={ko}")
for r in rows[1:]:
    if any(word in str(r[2]) for word in sys.argv[1:]):
        print("  ", r[0], "|", r[2], "|", r[3], "|", r[5], "|", r[6])
