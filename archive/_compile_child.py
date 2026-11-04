"""Child process: compile the whole VBA project of an .xlsm and print COMPILED_OK or COMPILE_FAILED."""
import sys
from pathlib import Path
import xlsm_ref as R

COMPILE_CONTROL_ID = 578   # VBE menu "Debug > Compile VBAProject" / "Débogage > Compiler VBAProject"

def main(path: str) -> None:
    with R.excel_session() as app:
        print(f"child excel pid {R.excel_pid(app)}", file=sys.stderr, flush=True)
        app.AutomationSecurity = R.MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        wb = app.Workbooks.Open(str(Path(path).resolve()))
        app.VBE.ActiveVBProject = wb.VBProject
        ctrl = app.VBE.CommandBars.FindControl(1, COMPILE_CONTROL_ID)
        print(f"before: enabled={ctrl.Enabled} caption={ctrl.Caption!r}", file=sys.stderr, flush=True)
        ctrl.Execute()
        enabled_after = ctrl.Enabled
        print(f"after: enabled={enabled_after}", file=sys.stderr, flush=True)
        print("COMPILED_OK" if not enabled_after else "COMPILE_FAILED", flush=True)
        del ctrl
        wb.Close(False)
        del wb

if __name__ == "__main__":
    main(sys.argv[1])
