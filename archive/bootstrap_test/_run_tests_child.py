"""Child: open the test workbook with macros enabled, run the VBA tests, save, print a JSON report."""
import json
import sys
from pathlib import Path

import xlsm_ref as R


def main(path: str, network: bool) -> None:
    with R.excel_session() as app:
        print(f"child excel pid {R.excel_pid(app)}", file=sys.stderr, flush=True)
        app.AutomationSecurity = R.MSO_AUTOMATION_SECURITY_LOW
        wb = app.Workbooks.Open(str(Path(path).resolve()))
        summary = app.Run(f"'{wb.Name}'!modTests.RunWithCoverage", network)   # positional argument
        rows = wb.Worksheets("TestResults").UsedRange.Value
        hits = wb.Worksheets("CoverageHits").UsedRange.Value
        wb.Save()
        wb.Close(False)
        del wb
    ko = [list(r[:7]) for r in rows[1:] if r[6] != "OK"]
    hit_ids = [r[0] for r in hits[1:] if r[0]]
    total = len(rows) - 1
    print(json.dumps({"summary": summary, "total": total, "ko": ko, "hits": hit_ids}, default=str))


if __name__ == "__main__":
    main(sys.argv[1], "--network" in sys.argv[2:])
