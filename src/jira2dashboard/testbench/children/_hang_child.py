"""Child that starts Excel, prints its PID, then hangs: tests the timeout cleanup."""
import sys
import time
from jira2dashboard.excel.com import excel_pid, excel_session

with excel_session() as app:
    print(f"child excel pid {excel_pid(app)}", file=sys.stderr, flush=True)
    time.sleep(300)
