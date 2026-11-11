"""Keeps the VBA log of an Excel scenario in its own file, next to the log of the workbook.

The workbook of the Excel scenarios writes `logs\\jira2dashboard.log` in its own folder (`out\\inter\\logs` for the test workbook, as
`out\\release\\logs` for the delivered one), and one log serves every scenario of a module. `mark` notes where the log ends before a
scenario, `keep` copies what the scenario added to `logs\\excel-test_<scenario>.log` in the same folder: a path relative to the
workbook, never a fixed one (the log of the VBA holds no token, URL or answer of Jira by design). `build` and `vba-tests` empty
`out\\inter`, logs included. Never raises: a log that cannot be kept must not fail a test.
"""
from __future__ import annotations

import re
from pathlib import Path

LOG_NAME = "jira2dashboard.log"


def file_name(scenario: str) -> str:
    """A name for a Windows file out of the name of a test (brackets and signs of the parametrize are dropped)."""
    return "excel-test_" + re.sub(r"[^A-Za-z0-9_.-]+", "_", scenario).strip("_")[:110] + ".log"


def mark(workbook_folder: str | Path) -> int:
    """Where the log of the workbook ends now (0 when there is none yet)."""
    try:
        return (Path(workbook_folder) / "logs" / LOG_NAME).stat().st_size
    except OSError:
        return 0


def keep(workbook_folder: str | Path, started_at: int, scenario: str) -> Path | None:
    """Copies the lines written since `started_at` to `<workbook folder>\\logs\\excel-test_<scenario>.log`; its path, None when
    there is nothing to keep."""
    try:
        logs = Path(workbook_folder) / "logs"
        with open(logs / LOG_NAME, "rb") as handle:
            if started_at > (logs / LOG_NAME).stat().st_size:       # the log was rotated (above 1 MB): take it whole
                started_at = 0
            handle.seek(started_at)
            data = handle.read()
        if not data.strip():
            return None
        target = logs / file_name(scenario)
        target.write_bytes(data)
        return target
    except OSError:
        return None
