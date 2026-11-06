"""What the code zip of a release holds (scripts/release.ps1 builds it with `git archive`).

`.gitattributes` marks the files for Claude, the old code and the local folders `export-ignore`: `git archive` leaves
them out, and so do the "Source code" archives GitHub adds to every release. The rest of the repo must stay in, so
that the client can build the workbook from the zip.
"""
import subprocess

import pytest

from jira2dashboard.paths import REPO_ROOT

LEFT_OUT = ("CLAUDE.md", "claude-code-handoff.md", "archive", "archive/old_probe.py", ".claude/settings.local.json",
            ".venv/Scripts/python.exe")
KEPT = ("README.md", "pyproject.toml", "requirements.txt", "scripts/run.ps1", "scripts/setup.ps1", "scripts/release.ps1",
        "src/jira2dashboard/build_workbook.py", "vba/modActions.bas", "params/config.md", "docs/rovo-guide.md",
        "tests/unit/test_checks.py", "rovo/rovo-guide.md")


def export_ignored(path: str) -> bool:
    out = subprocess.run(["git", "check-attr", "export-ignore", "--", path], cwd=REPO_ROOT, capture_output=True,
                         text=True, check=True).stdout
    return out.strip().endswith(": set")


@pytest.mark.parametrize("path", LEFT_OUT)
def test_the_release_zip_leaves_out(path):
    assert export_ignored(path)


@pytest.mark.parametrize("path", KEPT)
def test_the_release_zip_keeps(path):
    assert not export_ignored(path)


def release_notes_template() -> str:
    script = (REPO_ROOT / "scripts" / "release.ps1").read_text(encoding="utf-8")
    return script.split('$notes = @"', 1)[1].split('"@', 1)[0]


@pytest.mark.parametrize("algorithm", ("SHA256", "MD5"))
def test_the_release_script_hashes_both_files(algorithm):
    script = (REPO_ROOT / "scripts" / "release.ps1").read_text(encoding="utf-8")
    for file in ("$zip", "$xlsm"):
        assert f"Get-FileHash {file} -Algorithm {algorithm}" in script


def test_the_release_notes_give_the_md5_and_the_size_for_allow_lists():
    notes = release_notes_template()
    assert "MD5" in notes and "$xlsmMd5" in notes and "$zipMd5" in notes
    assert "$xlsmSize" in notes
