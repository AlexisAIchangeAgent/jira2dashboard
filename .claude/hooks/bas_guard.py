"""PostToolUse hook: after an Edit or Write of vba/*.bas, run the fast VBA guard tests.

Reads the hook JSON on stdin. For any other file it does nothing. For a .bas of vba/ it runs the pure tests that
check the VBA modules (ASCII, LF, Cov markers, no secret passed to modLog, no parameter literal), then:
- tests green: adds a reminder to Claude's context (the MD5 of the modules is no longer tracked);
- tests red: returns decision "block" with the end of the pytest output, so Claude fixes the module first.
"""
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
TESTS = ["tests/unit/test_run_vba_tests_pure.py", "tests/unit/test_vba_log_guard.py",
         "tests/unit/test_parameters_not_in_code.py"]


def edited_bas(payload: dict) -> Path | None:
    raw = (payload.get("tool_input") or {}).get("file_path") or (payload.get("tool_response") or {}).get("filePath")
    if not raw:
        return None
    path = Path(raw).resolve()
    return path if path.suffix.lower() == ".bas" and path.parent == ROOT / "vba" else None


def main() -> int:
    try:
        payload = json.loads(sys.stdin.read() or "{}")
    except json.JSONDecodeError:
        return 0
    module = edited_bas(payload)
    if module is None:
        return 0
    python = ROOT / ".venv" / "Scripts" / "python.exe"
    done = subprocess.run([str(python), "-m", "pytest", "-q", "-p", "no:cacheprovider", *TESTS], cwd=ROOT,
                          capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=110)
    tail = "\n".join((done.stdout + done.stderr).strip().splitlines()[-25:])
    if done.returncode != 0:
        print(json.dumps({"decision": "block",
                          "reason": f"VBA guard tests failed after editing {module.name}:\n{tail}\n"
                                    "Fix the module before going on."}))
        return 0
    summary = tail.splitlines()[-1] if tail else "passed"
    print(json.dumps({"hookSpecificOutput": {
        "hookEventName": "PostToolUse",
        "additionalContext": f"VBA guard tests after editing {module.name}: {summary}."}}))
    return 0


if __name__ == "__main__":
    sys.exit(main())
