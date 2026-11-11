"""Keeps the logs of each Excel test in a file of its own, in the folder logs of the test workbooks.

The test workbooks live in out\\inter, so their logs go to out\\inter\\logs (relative to the workbook, as the delivered one
writes out\\release\\logs): jira2dashboard.log (the VBA of every workbook of the tests), vba-tests.log (the VBA test
runner). One file serves every test. `mark` notes the size of each log before a test; `keep` writes
out\\inter\\logs\\excel-test_<test>.log after it: a header (test, result, time), what each log got during the test, then
the logs the test wrote in its temporary folder (python.log of the Python path). One file per test, even with no line,
so that each Excel test has its log. `clean` (build, vba-tests) keeps out\\inter\\logs. The logs of the VBA hold no token,
URL or answer of Jira by design; python.log never the token. Never raises: a log that cannot be kept must not fail a test.
"""
from __future__ import annotations

import re
from pathlib import Path

LOG_NAME = "jira2dashboard.log"
KEPT_PREFIX = "excel-test_"


def file_name(scenario: str) -> str:
    """A name for a Windows file out of the name of a test (brackets and signs of the parametrize are dropped)."""
    return KEPT_PREFIX + re.sub(r"[^A-Za-z0-9_.-]+", "_", scenario).strip("_")[:110] + ".log"


def _logs(folder: Path) -> list[Path]:
    try:
        return sorted(p for p in folder.glob("*.log") if not p.name.startswith(KEPT_PREFIX))
    except OSError:
        return []


def mark(logs_folder: str | Path) -> dict[str, int]:
    """The size of each log of the folder now (the kept files of the tests apart); empty when there is none yet."""
    sizes = {}
    for path in _logs(Path(logs_folder)):
        try:
            sizes[path.name] = path.stat().st_size
        except OSError:
            continue
    return sizes


def added(logs_folder: str | Path, marks: dict[str, int]) -> list[tuple[str, str]]:
    """(name, text) of what each log got since `marks`; a log smaller than its mark was rotated (above 1 MB): taken whole."""
    out = []
    for path in _logs(Path(logs_folder)):
        try:
            start = marks.get(path.name, 0)
            if start > path.stat().st_size:
                start = 0
            with open(path, "rb") as handle:
                handle.seek(start)
                data = handle.read()
        except OSError:
            continue
        if data.strip():
            out.append((path.name, data.decode("utf-8", "replace").replace("\r\n", "\n").rstrip("\n")))
    return out


def keep(logs_folder: str | Path, marks: dict[str, int], scenario: str, header: str,
         others: list[Path] | tuple[Path, ...] = ()) -> Path | None:
    """Writes <logs folder>\\excel-test_<scenario>.log; its path, None when it cannot be written."""
    parts = [header]
    for name, text in added(logs_folder, marks):
        parts += [f"--- {name}: the lines of this test", text]
    for path in others:
        try:
            text = path.read_text(encoding="utf-8", errors="replace").rstrip("\n")
        except OSError:
            continue
        parts += [f"--- {path.name}: written in the temporary folder of the test", text]
    if len(parts) == 1:
        parts.append("--- no log line during this test")
    try:
        target = Path(logs_folder) / file_name(scenario)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text("\n".join(parts) + "\n", encoding="utf-8", newline="\r\n")
        return target
    except OSError:
        return None
