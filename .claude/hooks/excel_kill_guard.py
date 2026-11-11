"""PreToolUse hook (Bash, PowerShell): ask before a command that stops Excel by its name.

CLAUDE.md: never kill an EXCEL.EXE the scripts did not start; AL may have their own Excel open (the one of
manual-test.ps1, or their work). "taskkill /IM EXCEL.EXE", "Stop-Process -Name EXCEL", "pkill excel" stop every Excel
at once. Only a command whose every kill targets processes by a PID number (/PID n, -Id n) and by no other way passes:
that is the documented way, after checking the start time (and the title "Microsoft Visual Basic" of an Excel hung on a
VBA error). The command is cut at ; & | && || and line breaks, and each part that kills must give a PID number. Anywhere
in the command, a name or filter (/IM, /FI, -Name, -ProcessName, -InputObject), a pipe into a kill, .Kill(), wmic, a CIM
Terminate or a Get-Process of excel makes it ask, even next to a PID. A guard against a mistake, not a sandbox: it reads
the text of the command, it does not parse the shell. The hook asks, it does not refuse: AL decides.
"""
import json
import re
import sys

KILL = re.compile(r"\b(taskkill|tskill|stop-process|spps|pkill|killall|kill|wmic|invoke-cimmethod|terminate)\b|\.kill\(",
                  re.IGNORECASE)
EXCEL = re.compile(r"excel", re.IGNORECASE)
ONE_PROCESS = re.compile(r"(/|-{1,2})(pid|id)\s*[:=]?\s*\d+", re.IGNORECASE)
BY_NAME = re.compile(r"(^|\s)(/{1,2}|-)(im|fi)\b|-(process)?name\b|-inputobject\b|"
                     r"\|\s*(stop-process|spps|kill|taskkill|foreach|%)\b|\.kill\(|"
                     r"\b(wmic|invoke-cimmethod|pkill|killall|tskill)\b|\bget-process\b[^|;&\n]*excel", re.IGNORECASE)
PARTS = re.compile(r"&&|\|\||[;&|\n]")
REASON = ("This command stops Excel by its name, so every EXCEL.EXE, AL's own too (CLAUDE.md: never kill an Excel the "
          "scripts did not start). Stop one process instead: find the Excel of the run by its start time "
          "(Get-Process EXCEL | Select Id, StartTime, MainWindowTitle), then taskkill /PID <id>.")


def kills_excel_by_name(command: str) -> bool:
    """True for a command that stops Excel otherwise than by PID numbers only."""
    if not (KILL.search(command) and EXCEL.search(command)):
        return False
    if BY_NAME.search(command):
        return True
    return any(KILL.search(part) and not ONE_PROCESS.search(part) for part in PARTS.split(command))


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
