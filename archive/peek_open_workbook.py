"""Read-only look at the workbook open in the user's Excel (no change, no save)."""
import sys
import pythoncom
import pywintypes
import win32com.client as win32

PATH = r"C:\dev\xlsm-check\out\vba_tests_boutons.xlsm"
pythoncom.CoInitialize()
try:
    wb = win32.GetObject(PATH)                       # attaches to the already open workbook
    app = wb.Application
    print("Excel ready:", app.Ready, "| StatusBar:", app.StatusBar, "| Calculation:", app.Calculation)
    print("AutomationSecurity:", app.AutomationSecurity)
    used = wb.Worksheets("TestResults").UsedRange
    print("UsedRange:", used.Address, "| last row in A:", wb.Worksheets("TestResults").Cells(65536, 1).End(-4162).Row)
    rows = wb.Worksheets("TestResults").Range("A1:H140").Value
    rows = tuple(r for r in rows if r[0] is not None)
    print("TestResults rows (with header):", len(rows))
    if isinstance(rows, tuple):
        areas = {}
        for r in rows[1:]:
            areas.setdefault(r[1], []).append(r[6])
        for a, res in areas.items():
            print(f"  {a:11} {len(res):3} rows, KO={sum(1 for x in res if x != 'OK')}")
        for r in rows[1:]:
            if r[1] in ("net", "runner") or r[6] != "OK":
                print("  ", " | ".join(str(x)[:60] for x in r[:7]))
    del wb, app
except pywintypes.com_error as e:
    print("COM error:", e.hresult, e.excepinfo[2] if e.excepinfo else e.strerror)
    sys.exit(1)
finally:
    pythoncom.CoUninitialize()
