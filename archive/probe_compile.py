"""Probe: run _compile_child.py on a good and a broken file, with timeout and Excel cleanup."""
import re
import subprocess
import sys
import time
from pathlib import Path
import xlsm_ref as R

HERE = Path(__file__).parent

def compile_in_child(xlsm: Path, timeout_s: int = 45) -> str:
    t0 = time.time()
    try:
        res = subprocess.run([sys.executable, str(HERE / "_compile_child.py"), str(xlsm)],
                             capture_output=True, text=True, timeout=timeout_s)
        tail = res.stderr.strip().splitlines()[-3:]
        return f"{res.stdout.strip()!r} in {time.time() - t0:.1f}s exit={res.returncode} stderr={tail}"
    except subprocess.TimeoutExpired as err:
        stderr = err.stderr.decode() if isinstance(err.stderr, bytes) else (err.stderr or "")
        m = re.search(r"child excel pid (\d+)", stderr)
        if m:
            R._kill_if_still_excel(int(m.group(1)), grace_s=0)
        return f"TIMEOUT after {time.time() - t0:.1f}s, killed child Excel pid {m.group(1) if m else '?'}; stderr={stderr.strip().splitlines()[-2:]}"

for name in ["check.xlsm", "compile_error.xlsm"]:
    print(f"{name}: {compile_in_child(HERE / 'out' / name)}")
