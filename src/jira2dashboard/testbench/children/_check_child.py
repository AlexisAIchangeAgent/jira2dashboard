"""Internal: run by testbench/run_vba_tests.py and cli/check_machine.py in a process of its own.

Opens an .xlsm in a new Excel, compiles its VBA, optionally runs modSelfTest.Ping, and prints COMPILED_OK or
COMPILE_FAILED (and the answer of Ping). Not meant to be run by hand.
"""
import sys
from pathlib import Path

from jira2dashboard.cli.args import parser
from jira2dashboard.excel.com import (
    MSO_AUTOMATION_SECURITY_FORCE_DISABLE, MSO_AUTOMATION_SECURITY_LOW,
    compile_vba_project, excel_pid, excel_session, run_macro,
)

def main(path: str, run_ping: bool) -> None:
    with excel_session() as app:
        print(f"child excel pid {excel_pid(app)}", file=sys.stderr, flush=True)
        app.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW if run_ping else MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        wb = app.Workbooks.Open(str(Path(path).resolve()))
        ok = compile_vba_project(app, wb)
        result = "COMPILED_OK" if ok else "COMPILE_FAILED"
        if ok and run_ping:
            result += " " + str(run_macro(app, wb, "modSelfTest.Ping"))
        print(result, flush=True)
        wb.Close(False)
        del wb

if __name__ == "__main__":
    p = parser("_check_child.py", __doc__)
    p.add_argument("workbook", help="the .xlsm to compile")
    p.add_argument("--ping", action="store_true", help="also run modSelfTest.Ping (macros enabled)")
    args = p.parse_args()
    main(args.workbook, run_ping=args.ping)
