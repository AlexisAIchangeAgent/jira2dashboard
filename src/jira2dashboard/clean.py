"""Clean the folder out before a build: out\\release and out\\manual_test stay, out\\inter is emptied, nothing else
stays at the root.

build and vba-tests run it first. out\\release holds what is delivered (the workbook of `build`, the release packages
of scripts\\release.ps1) and out\\manual_test the copy of scripts\\manual-test.ps1 (jira2dashboard4manualtest.xlsm,
filled for the fake Jira): never touched. out\\inter holds the rest (test workbooks, exports, logs): emptied, except
the tool caches that pyproject.toml puts there (.coverage, .pytest_cache). Anything else at the root of out is a
leftover of the old layout and goes. A file that a program still holds (Windows refuses to delete it) is kept and
reported, never forced.
"""
import shutil
import sys
from pathlib import Path

from jira2dashboard.cli.args import parser
from jira2dashboard.paths import INTER_DIR, MANUAL_TEST_DIR, OUT_DIR, RELEASE_DIR

TOOL_CACHES = (".coverage", ".pytest_cache")
FOLDERS = (RELEASE_DIR.name, INTER_DIR.name, MANUAL_TEST_DIR.name)


def entries(folder: Path) -> list[Path]:
    return list(folder.iterdir()) if folder.is_dir() else []


def intermediates(out_dir: Path) -> list[Path]:
    leftovers = [entry for entry in entries(out_dir) if entry.name not in FOLDERS]
    inter = [entry for entry in entries(out_dir / INTER_DIR.name) if entry.name not in TOOL_CACHES]
    return sorted(leftovers + inter)


def clean(out_dir: Path = OUT_DIR) -> tuple[list[Path], list[Path]]:
    """Remove the intermediates of out_dir. Returns (removed, kept because in use)."""
    removed, in_use = [], []
    for entry in intermediates(out_dir):
        try:
            if entry.is_dir() and not entry.is_symlink():
                shutil.rmtree(entry)
            else:
                entry.unlink()
            removed.append(entry)
        except OSError:
            in_use.append(entry)
    return removed, in_use


def parse(argv: list[str] | None = None):
    return parser("scripts\\run.ps1 clean", __doc__, "example:\n  scripts\\run.ps1 clean").parse_args(argv)


def main(out_dir: Path = OUT_DIR) -> int:
    removed, in_use = clean(out_dir)
    print(f"clean: {len(removed)} intermediate(s) removed from {out_dir}")
    for entry in in_use:
        print(f"clean: kept {entry.name} (in use)")
    return 0


if __name__ == "__main__":
    parse()
    sys.exit(main())
