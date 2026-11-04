"""Child process: open an .xlsm with macros enabled and print the result of Ping."""
import logging
import sys
from pathlib import Path

import xlsm_ref as R

logging.basicConfig(level=logging.INFO, stream=sys.stderr)

def main(path: str) -> None:
    with R.excel_session() as app:
        print(f"child excel pid {R.excel_pid(app)}", file=sys.stderr, flush=True)
        app.AutomationSecurity = R.MSO_AUTOMATION_SECURITY_LOW
        wb = app.Workbooks.Open(str(Path(path).resolve()))
        print(R.run_macro(app, wb, "modSelfTest.Ping"), flush=True)
        wb.Close(False)

if __name__ == "__main__":
    main(sys.argv[1])
