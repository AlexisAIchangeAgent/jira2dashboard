"""Keeps the VBA log of an Excel scenario, so that it can be read after the run.

The workbook of the Excel scenarios writes `logs\\jira2dashboard.log` next to itself, in `out\\inter`, which `build` and `vba-tests`
empty, and one log serves every scenario of a module. `mark` notes where the log ends before a scenario, `keep` copies what the
scenario added to `scripts\\logs\\excel-tests\\<scenario>.log` (the log of the VBA holds no token, URL or
answer of Jira by design). Never raises: a log that cannot be kept must not fail a test.
"""
from __future__ import annotations

import re
from pathlib import Path

from jira2dashboard.paths import REPO_ROOT

LOG_NAME = "jira2dashboard.log"


def kept_folder(logs_folder: str | Path | None = None) -> Path:
    """`scripts\\logs\\excel-tests`, the folder of the logs of the scripts (a test points `logs_folder` elsewhere)."""
    return Path(logs_folder or REPO_ROOT / "scripts" / "logs") / "excel-tests"


def file_name(scenario: str) -> str:
    """A name for a Windows file out of the name of a test (brackets and signs of the parametrize are dropped)."""
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", scenario).strip("_")[:120] + ".log"


def mark(workbook_folder: str | Path) -> int:
    """Where the log of the workbook ends now (0 when there is none yet)."""
    try:
        return (Path(workbook_folder) / "logs" / LOG_NAME).stat().st_size
    except OSError:
        return 0


def keep(workbook_folder: str | Path, started_at: int, scenario: str, logs_folder: str | Path | None = None) -> Path | None:
    """Copies the lines written since `started_at`; the path of the copy, None when there is nothing to keep."""
    try:
        source = Path(workbook_folder) / "logs" / LOG_NAME
        with open(source, "rb") as handle:
            if started_at > source.stat().st_size:                  # the log was rotated (above 1 MB): take it whole
                started_at = 0
            handle.seek(started_at)
            data = handle.read()
        if not data.strip():
            return None
        target = kept_folder(logs_folder) / file_name(scenario)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)
        return target
    except OSError:
        return None
