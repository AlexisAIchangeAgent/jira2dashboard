"""Child: open the test workbook with macros enabled, run the VBA tests, save, print a JSON report.

    _run_tests_child.py <workbook> [--network] [--pac-url <url of a local proxy script, for T_ProxyScript>]
"""
import json
import sys
from pathlib import Path

from jira2dashboard.excel import com as R


def grid(value):
    """A range of one cell comes back as a scalar: make it a one-row table like any other."""
    return value if isinstance(value, tuple) else ((value,),)


def option(args: list[str], flag: str) -> str:
    return args[args.index(flag) + 1] if flag in args[:-1] else ""


def main(path: str, network: bool, pac_url: str) -> None:
    with R.excel_session() as app:
        print(f"child excel pid {R.excel_pid(app)}", file=sys.stderr, flush=True)
        app.AutomationSecurity = R.MSO_AUTOMATION_SECURITY_LOW
        wb = app.Workbooks.Open(str(Path(path).resolve()))
        app.Run(f"'{wb.Name}'!modTests.SetPacUrl", pac_url)                   # "" = T_ProxyScript checks nothing
        summary = app.Run(f"'{wb.Name}'!modTests.RunWithCoverage", network)   # positional argument
        rows = grid(wb.Worksheets("TestResults").UsedRange.Value)
        hits = grid(wb.Worksheets("CoverageHits").UsedRange.Value)
        wb.Save()
        wb.Close(False)
        del wb
    ko = [list(r[:7]) for r in rows[1:] if r[6] != "OK"]
    hit_ids = [r[0] for r in hits[1:] if r[0]]
    total = len(rows) - 1
    print(json.dumps({"summary": summary, "total": total, "ko": ko, "hits": hit_ids}, default=str))


if __name__ == "__main__":
    main(sys.argv[1], "--network" in sys.argv[2:], option(sys.argv[2:], "--pac-url"))
