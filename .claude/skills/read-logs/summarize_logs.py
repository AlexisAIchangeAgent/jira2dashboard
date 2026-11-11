"""Summary of the logs of jira2dashboard, for the skill read-logs: what failed (with its support code), what was slow, what
was retried, what did not end. Standard library only; reads, never writes.

    .venv\\Scripts\\python.exe .claude\\skills\\read-logs\\summarize_logs.py [FILE_OR_FOLDER ...]

Without an argument: out\\inter\\logs, out\\manual_test\\logs, out\\release\\logs and scripts\\logs of the repository.
Formats read: the VBA log ("2026-11-11 20:24:18.906 KIND text"), python.log ("2026-11-11 20:24:18,906 LEVEL text"), the
kept log of an Excel test (first line "# test | result | time | date") and the transcripts of scripts\\logs (pytest lines).
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
DEFAULT = [ROOT / "out" / "inter" / "logs", ROOT / "out" / "manual_test" / "logs", ROOT / "out" / "release" / "logs",
           ROOT / "scripts" / "logs"]
LINE = re.compile(r"^(\d{4}-\d\d-\d\d \d\d:\d\d:\d\d[.,]\d{3}) (\w+) (.*)$")
TIMED = re.compile(r"^(.*) \| (\d+\.\d{3}) s$")
CODE = re.compile(r"J2D-[A-Z]+-[0-9A-F]+")
PYTEST = re.compile(r"(^FAILED |^ERROR |=+ .*(passed|failed|error).* in [\d.]+s)")
SLOWEST = 5


def files(targets: list[Path]) -> list[Path]:
    out = []
    for target in targets:
        if target.is_dir():
            out += sorted(target.glob("*.log"))
        elif target.is_file():
            out.append(target)
    return out


def summarize(name: str, lines: list[str]) -> list[str]:
    """The lines worth reading in one log file."""
    out: list[str] = []
    if lines and lines[0].startswith("# "):                          # the kept log of an Excel test
        header = lines[0][2:].split(" | ")
        if len(header) >= 2 and header[1] != "passed":
            out.append(f"TEST {header[1].upper()}: {header[0]} ({' | '.join(header[2:])})")
    steps: list[tuple[float, str]] = []
    running = ""                                                       # the RUN line of a run without its RUN end yet
    for line in lines:
        if PYTEST.search(line):
            out.append(f"pytest: {line.strip()}")
            continue
        match = LINE.match(line)
        if not match:
            continue
        stamp, kind, text = match.groups()
        if kind == "RUN" and (" | Excel " in text or text.startswith("extract |")):
            if running:
                out.append(f"NOT ENDED (crash, or Excel closed during the run): {running}")
            running = f"{stamp} RUN {text}"
            out.append(running)
        elif kind == "RUN" and text.startswith("end "):
            running = ""
            out.append(f"{stamp} {kind} {text}")
        elif kind == "RUN" or kind in ("ERROR", "RETRY", "WARNING"):
            out.append(f"{stamp} {kind} {text}")
        elif kind in ("END", "INFO") and TIMED.match(text):
            what, seconds = TIMED.match(text).groups()
            steps.append((float(seconds), f"{stamp} {what.removeprefix('END ')}"))
    if running:
        out.append(f"NOT ENDED (crash, still running, or Excel closed during the run): {running}")
    steps.sort(reverse=True)
    out += [f"slow: {seconds:.3f} s  {what}" for seconds, what in steps[:SLOWEST] if seconds >= 0.5]
    return [f"== {name}"] + out if out else []


def codes(report: list[str]) -> list[str]:
    return sorted({code for line in report for code in CODE.findall(line)})


def main(argv: list[str]) -> int:
    targets = [Path(arg) for arg in argv] or DEFAULT
    report: list[str] = []
    for path in files(targets):
        try:
            text = path.read_text(encoding="utf-8-sig", errors="replace")
        except OSError as err:
            report.append(f"== {path}: not read ({err})")
            continue
        report += summarize(str(path), text.splitlines())
    print("\n".join(report) if report else "Nothing to report: no failure, no error, no retry, no slow step.")
    found = codes(report)
    if found:
        print("Support codes seen: " + ", ".join(found) + " (meaning: python -m jira2dashboard.error_codes)")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
