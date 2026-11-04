"""Read-only: print the first cells of TestResults and CoverageHits in the user's open workbook."""
import pythoncom
import win32com.client as win32

pythoncom.CoInitialize()
wb = win32.GetObject(r"C:\dev\xlsm-check\out\vba_tests_boutons.xlsm")
for name, addr in [("TestResults", "A1:H4"), ("TestResults", "A128:H131"), ("CoverageHits", "A1:B3")]:
    print(f"--- {name}!{addr}")
    for r in wb.Worksheets(name).Range(addr).Value:
        print("  ", r)
ws = wb.Worksheets("TestResults")
print("Shapes:", [(ws.Shapes(i).Name, ws.Shapes(i).OnAction) for i in range(1, ws.Shapes.Count + 1)])
del ws, wb
pythoncom.CoUninitialize()
