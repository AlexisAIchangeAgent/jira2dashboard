"""PreToolUse hook (Bash, PowerShell): ask before a command that stops Excel by its name.

CLAUDE.md: never kill an EXCEL.EXE the scripts did not start; AL may have their own Excel open (the one of
manual-test.ps1, or their work). "taskkill /IM EXCEL.EXE", "Stop-Process -Name EXCEL", "pkill excel" stop every Excel
at once. A command that names one process (/PID n, -Id n) passes: that is the documented way, after checking the start
time (and the title "Microsoft Visual Basic" of an Excel hung on a VBA error). The hook asks, it does not refuse: AL
decides.
"""
import json
import re
import sys

KILL = re.compile(r"\b(taskkill|stop-process|pkill|killall|kill)\b", re.IGNORECASE)
EXCEL = re.compile(r"excel", re.IGNORECASE)
ONE_PROCESS = re.compile(r"(/|-{1,2})(pid|id)\s*[:=]?\s*\d+", re.IGNORECASE)
REASON = ("This command stops Excel by its name, so every EXCEL.EXE, AL's own too (CLAUDE.md: never kill an Excel the "
          "scripts did not start). Stop one process instead: find the Excel of the run by its start time "
          "(Get-Process EXCEL | Select Id, StartTime, MainWindowTitle), then taskkill /PID <id>.")


def kills_excel_by_name(command: str) -> bool:
    """True for a command that stops Excel without naming one process."""
    return bool(KILL.search(command) and EXCEL.search(command) and not ONE_PROCESS.search(command))


def main() -> int:
    try:
        payload = json.loads(sys.stdin.read() or "{}")
    except json.JSONDecodeError:
        return 0
    command = str((payload.get("tool_input") or {}).get("command") or "")
    if kills_excel_by_name(command):
        print(json.dumps({"hookSpecificOutput": {"hookEventName": "PreToolUse", "permissionDecision": "ask",
                                                 "permissionDecisionReason": REASON}}))
    return 0


if __name__ == "__main__":
    sys.exit(main())
