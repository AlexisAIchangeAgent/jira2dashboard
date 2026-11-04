"""Probe: what happens when the VBA project has a compile error and Ping is run."""
import subprocess
import sys
import time
from pathlib import Path
import xlsm_ref as R

OUT = Path(__file__).with_name("out") / "compile_error.xlsm"
BAD = "Option Explicit\r\n\r\nPublic Sub Broken()\r\n    x = 1\r\nEnd Sub\r\n"   # x not declared

def build() -> None:
    with R.excel_session() as app:
        wb = R.new_workbook(app)
        R.add_or_replace_module(wb, "modBad", BAD)
        R.add_or_replace_module(wb, "modSelfTest", R.normalize_vba(R.MOD_SELF_TEST))
        R.save_as_xlsm(wb, OUT, overwrite=True)
        wb.Close(False)
        del wb

def run_child(timeout_s: int) -> None:
    t0 = time.time()
    try:
        out = R.run_ping_in_child(OUT, timeout_s=timeout_s)
        print(f"child returned in {time.time() - t0:.1f}s: {out!r}")
    except subprocess.TimeoutExpired as err:
        print(f"TIMEOUT after {time.time() - t0:.1f}s; child stderr: {err.stderr!r}")
    except subprocess.CalledProcessError as err:
        tail = (err.stderr or "").strip().splitlines()[-3:]
        print(f"child failed in {time.time() - t0:.1f}s, exit {err.returncode}: {tail}")

if __name__ == "__main__":
    build()
    print("built", OUT)
    run_child(int(sys.argv[1]) if len(sys.argv) > 1 else 40)
