"""Internal: run by cli/check_machine.py to test the clean-up of a hung child.

Starts Excel, prints its pid, then hangs for 300 s, so that the caller's timeout must kill both. Not meant to be run
by hand.
"""
import sys
import time

from jira2dashboard.cli.args import parser
from jira2dashboard.excel.com import excel_pid, excel_session

if __name__ == "__main__":
    p = parser("_hang_child.py", __doc__)
    p.add_argument("workbook", nargs="?", help="ignored: run_child passes a workbook to every child")
    p.parse_args()
    with excel_session() as app:
        print(f"child excel pid {excel_pid(app)}", file=sys.stderr, flush=True)
        time.sleep(300)
