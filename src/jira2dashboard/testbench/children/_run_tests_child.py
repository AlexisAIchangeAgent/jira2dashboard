"""Internal: run by testbench/run_vba_tests.py in a process of its own.

Opens the test workbook with macros enabled in a new Excel, runs the VBA tests (modTests.RunWithCoverage), saves it,
and prints a JSON report: summary, total, the KO rows and the coverage markers hit. Not meant to be run by hand.
"""
import json
import sys
from pathlib import Path

from jira2dashboard.cli.args import parser
from jira2dashboard.excel import com as R


def grid(value):
    """A range of one cell comes back as a scalar: make it a one-row table like any other."""
    return value if isinstance(value, tuple) else ((value,),)


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
    p = parser("_run_tests_child.py", __doc__)
    p.add_argument("workbook", help="the test workbook (.xlsm) built by run_vba_tests")
    p.add_argument("--network", action="store_true", help="also run the network tests N_")
    p.add_argument("--pac-url", default="", help="address of a local proxy script, for T_ProxyScript")
    args = p.parse_args()
    main(args.workbook, args.network, args.pac_url)
