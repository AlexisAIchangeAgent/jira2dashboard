"""Child: open an .xlsm, compile its VBA, optionally run Ping, print the result."""
import sys
from pathlib import Path

from xlsm_ref import (   # adapt to the module that holds the helpers
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
    main(sys.argv[1], run_ping="--ping" in sys.argv[2:])
