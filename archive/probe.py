"""Probe: datetime round trip, Worksheets.Add placement, Quit timing."""
import datetime as dt
import time
import pythoncom
import win32api, win32con
import xlsm_ref as R

def alive(pid):
    try:
        h = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION, False, pid)
        win32api.CloseHandle(h)
        return True
    except Exception:
        return False

def names(wb):
    return [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]

pythoncom.CoInitialize()
app = R.start_excel()
pid = R.excel_pid(app)
wb = R.new_workbook(app)
ws = wb.Worksheets(1); ws.Name = "Data"

print("== datetime ==")
for label, value in [("naive", dt.datetime(2026, 11, 3, 0, 0)),
                     ("utc-aware", dt.datetime(2026, 11, 3, tzinfo=dt.timezone.utc)),
                     ("date", dt.date(2026, 11, 3))]:
    try:
        ws.Range("A1").Value = value
        back = ws.Range("A1").Value
        print(f"{label:10} wrote={value!r} read={back!r} Value2={ws.Range('A1').Value2!r} Text={ws.Range('A1').Text!r}")
    except Exception as e:
        print(f"{label:10} ERROR {e!r}")

print("== Worksheets.Add placement ==")
variants = {
    "After=last (named)": lambda last: wb.Worksheets.Add(After=last),
    "(None, last)": lambda last: wb.Worksheets.Add(None, last),
    "(pythoncom.Empty, last)": lambda last: wb.Worksheets.Add(pythoncom.Empty, last),
    "(pythoncom.Missing, last)": lambda last: wb.Worksheets.Add(pythoncom.Missing, last),
    "Add() then Move(None, last)": lambda last: (lambda s: (s.Move(None, last), s)[1])(wb.Worksheets.Add()),
}
for i, (label, fn) in enumerate(variants.items(), 1):
    last = wb.Worksheets(wb.Worksheets.Count)
    try:
        s = fn(last)
        s = wb.ActiveSheet if s is None else s
        s.Name = f"V{i}"
        print(f"{label:30} -> {names(wb)}")
    except Exception as e:
        print(f"{label:30} ERROR {e!r}")

print("== Quit timing ==")
wb.Close(False)
app.Quit()
for ref in ("ws", "wb", "s", "last"):
    globals().pop(ref, None)
t0 = time.time()
print("alive right after Quit (app ref kept):", alive(pid))
del app
import gc; gc.collect()
while alive(pid) and time.time() - t0 < 15:
    time.sleep(0.25)
print(f"alive after releasing refs: {alive(pid)}  waited {time.time() - t0:.2f}s")
pythoncom.CoUninitialize()
