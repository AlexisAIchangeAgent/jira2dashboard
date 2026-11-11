"""The version of a workbook: written at build into its hidden name J2D_Version, read back on the RUN line of the logs.

"v1.0.0 commit 1a2b3c4, built 2026-11-11 19:40" when scripts\\release.ps1 builds it (--version), "commit 1a2b3c4, built ..."
otherwise, "with local changes" when the tree was not clean, "commit unknown" without git (the code from a zip). The
commit, not git describe: the tags of the releases are on GitHub and may be missing from a clone. The VBA reads it with
modLog.WorkbookVersion, the Python path with standalone.xlsx_read.defined_name.
"""
from __future__ import annotations

import datetime as dt
import subprocess
from pathlib import Path

from jira2dashboard.paths import REPO_ROOT

NAME = "J2D_Version"
UNKNOWN = "unknown"


def commit(root: Path = REPO_ROOT, run=subprocess.run) -> tuple[str, bool]:
    """(short hash of HEAD, the tree has changes); ("unknown", False) when git or the repository is missing."""
    def git(*args: str) -> str:
        return run(["git", "-C", str(root), *args], capture_output=True, text=True, check=True).stdout.strip()
    try:
        return git("rev-parse", "--short", "HEAD") or UNKNOWN, bool(git("status", "--porcelain"))
    except (OSError, subprocess.CalledProcessError):
        return UNKNOWN, False


def stamp(release: str | None, head: str, dirty: bool, built: dt.datetime) -> str:
    """The text of the name J2D_Version."""
    text = f"commit {head}" + (" with local changes" if dirty else "") + f", built {built:%Y-%m-%d %H:%M}"
    return f"{release} {text}" if release else text


def formula(text: str) -> str:
    """The RefersTo of the name: a text constant, its quotes doubled."""
    return '="' + text.replace('"', '""') + '"'
