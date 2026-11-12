"""What the "Source code" archives of a release hold, and what release.ps1 ships (the workbook only, no zip).

`.gitattributes` marks the files for Claude, the old code and the local folders `export-ignore`: `git archive` leaves
them out, and so do the "Source code" archives GitHub adds to every release (the only code a release holds). The rest
of the repo must stay in, so that the client can build the workbook from them.
"""
import subprocess

import pytest

from jira2dashboard.paths import DELIVERED_NAME, REPO_ROOT

LEFT_OUT = ("CLAUDE.md", "archive", "archive/old_probe.py", ".claude/settings.local.json", ".claude/settings.json", ".claude/hooks/bas_guard.py",
            ".venv/Scripts/python.exe")
KEPT = ("README.md", "pyproject.toml", "requirements.txt", "scripts/run.ps1", "scripts/run.cmd", "scripts/_launch.cmd", "scripts/setup.ps1", "scripts/release.ps1",
        "src/jira2dashboard/build_workbook.py", "vba/modActions.bas", "params/config.md", "docs/rovo-guide.md",
        "tests/unit/test_checks.py", "rovo/rovo-guide.md")


def export_ignored(path: str) -> bool:
    out = subprocess.run(["git", "check-attr", "export-ignore", "--", path], cwd=REPO_ROOT, capture_output=True,
                         text=True, check=True).stdout
    return out.strip().endswith(": set")


@pytest.mark.parametrize("path", LEFT_OUT)
def test_the_source_archives_leave_out(path):
    assert export_ignored(path)


@pytest.mark.parametrize("path", KEPT)
def test_the_source_archives_keep(path):
    assert not export_ignored(path)


def release_notes_template() -> str:
    script = (REPO_ROOT / "scripts" / "release.ps1").read_text(encoding="utf-8")
    return script.split('$notes = @"', 1)[1].split('"@', 1)[0]


@pytest.mark.parametrize("algorithm", ("SHA256", "MD5"))
def test_the_release_script_hashes_the_workbook(algorithm):
    script = (REPO_ROOT / "scripts" / "release.ps1").read_text(encoding="utf-8")
    assert f"Get-FileHash $xlsm -Algorithm {algorithm}" in script


def test_the_release_ships_the_workbook_only_no_zip_of_the_code():
    script = (REPO_ROOT / "scripts" / "release.ps1").read_text(encoding="utf-8")
    assert "gh release create $Version $xlsm --target" in script
    assert "git archive" not in script.split("#>", 1)[1] and "$zip" not in script


def test_every_release_ships_the_workbook_under_the_same_name_and_the_notes_say_why():
    script = (REPO_ROOT / "scripts" / "release.ps1").read_text(encoding="utf-8")
    assert f'Join-Path $dist "{DELIVERED_NAME}"' in script       # the name never carries the version
    assert f"out\\release\\{DELIVERED_NAME}" in script
    notes = release_notes_template()
    assert f"keep the file name {DELIVERED_NAME}" in notes and "same folder on every machine" in notes
    assert "replaces the file in place" in notes


def test_the_release_notes_give_the_md5_and_the_size_for_allow_lists():
    notes = release_notes_template()
    assert "MD5" in notes and "$xlsmMd5" in notes and ".zip" not in notes
    assert "$xlsmSize" in notes
