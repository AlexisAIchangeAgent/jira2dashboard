"""PostToolUse hook: after a change of vba/*.bas, run the fast VBA guard tests.

Reads the hook JSON on stdin.
- Edit or Write of a .bas of vba/: the tests run.
- Bash or PowerShell: the tests run when a .bas of vba/ changed since the last check (a Python edit script, a cp that
  restores a mutated module): the size and time of each module are kept in %TEMP%\\jira2dashboard\\bas_guard.json.
  The first command of a machine only records them.
- anything else: nothing.
The tests check the VBA modules (ASCII, LF, Cov markers, no secret passed to modLog, no parameter literal):
- green: a reminder in Claude's context;
- red: decision "block" with the end of the pytest output, so Claude fixes the module first.
"""
import json
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
VBA_DIR = ROOT / "vba"
STATE = Path(tempfile.gettempdir()) / "jira2dashboard" / "bas_guard.json"
TESTS = ["tests/unit/test_run_vba_tests_pure.py", "tests/unit/test_vba_log_guard.py",
         "tests/unit/test_parameters_not_in_code.py"]
SHELLS = ("Bash", "PowerShell")


def edited_bas(payload: dict, vba_dir: Path = VBA_DIR) -> Path | None:
    raw = (payload.get("tool_input") or {}).get("file_path") or (payload.get("tool_response") or {}).get("filePath")
    if not raw:
        return None
    path = Path(raw).resolve()
    return path if path.suffix.lower() == ".bas" and path.parent == vba_dir.resolve() else None


def fingerprint(vba_dir: Path = VBA_DIR) -> dict[str, list[int]]:
    """{module: [size, modification time in ns]} of each .bas of the folder."""
    out = {}
    for path in sorted(vba_dir.glob("*.bas")):
        try:
            stat = path.stat()
        except OSError:
            continue
        out[path.name] = [stat.st_size, stat.st_mtime_ns]
    return out


def changed_modules(vba_dir: Path = VBA_DIR, state: Path = STATE) -> list[str]:
    """The modules added, removed or changed since the last call (none at the first one); the state is updated."""
    now = fingerprint(vba_dir)
    try:
        before = json.loads(state.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        before = None
    try:
        state.parent.mkdir(parents=True, exist_ok=True)
        state.write_text(json.dumps(now), encoding="utf-8")
    except OSError:
        pass
    if not isinstance(before, dict):
        return []
    return sorted(name for name in set(now) | set(before) if now.get(name) != before.get(name))


def run_tests(label: str) -> dict:
    python = ROOT / ".venv" / "Scripts" / "python.exe"
    done = subprocess.run([str(python), "-m", "pytest", "-q", "-p", "no:cacheprovider", *TESTS], cwd=ROOT,
                          capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=110)
    tail = "\n".join((done.stdout + done.stderr).strip().splitlines()[-25:])
    if done.returncode != 0:
        return {"decision": "block",
                "reason": f"VBA guard tests failed after {label}:\n{tail}\nFix the module before going on."}
    summary = tail.splitlines()[-1] if tail else "passed"
    return {"hookSpecificOutput": {"hookEventName": "PostToolUse",
                                   "additionalContext": f"VBA guard tests after {label}: {summary}."}}


def main() -> int:
    try:
        payload = json.loads(sys.stdin.read() or "{}")
    except json.JSONDecodeError:
        return 0
    tool = payload.get("tool_name")
    if tool in SHELLS:
        changed = changed_modules()
        if not changed:
            return 0
        label = "a command changed " + ", ".join(changed)
    else:
        module = edited_bas(payload)
        if module is None:
            return 0
        changed_modules()                                  # the next command does not test the same edit again
        label = f"editing {module.name}"
    print(json.dumps(run_tests(label)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
