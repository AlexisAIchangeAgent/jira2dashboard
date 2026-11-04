"""Reproduce the network button path (clear, then RunWithCoverage(True)) on a copy, timed. No MsgBox."""
import shutil
import sys
import time
from pathlib import Path
import xlsm_ref as R

HERE = Path(__file__).parent
SRC = HERE / "out" / "vba_tests_boutons.xlsm"
COPY = HERE / "out" / "repro_copy.xlsm"
shutil.copyfile(SRC, COPY)

with R.excel_session() as app:
    print(f"child excel pid {R.excel_pid(app)}", file=sys.stderr, flush=True)
    app.AutomationSecurity = R.MSO_AUTOMATION_SECURITY_LOW
    wb = app.Workbooks.Open(str(COPY))
    app.Run(f"'{wb.Name}'!modButtons.BtnClearResults")
    for label, network in (("offline", False), ("offline + network", True)):
        app.Run(f"'{wb.Name}'!modButtons.BtnClearResults")
        t0 = time.time()
        summary = app.Run(f"'{wb.Name}'!modTests.RunWithCoverage", network)
        n = wb.Worksheets("TestResults").Cells(65536, 1).End(-4162).Row - 1
        print(f"{label:18} {time.time() - t0:5.1f} s | {summary} | rows written: {n}", flush=True)
    wb.Close(False)
    del wb
