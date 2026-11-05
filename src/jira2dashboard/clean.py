"""Clean the folder out before a build: the workbooks stay, the intermediate files go.

Usage:  scripts\\run.ps1 clean   (build and vba-tests also clean first)

Kept, at the top of out: the workbooks (*.xlsm, which also keeps Excel's lock file ~$name.xlsm of an open workbook)
and the tool caches that pyproject.toml puts there (.coverage, .pytest_cache). Everything else goes: logs, exports,
temporary folders, old copies. A file that a program still holds (Windows refuses to delete it) is kept and reported,
never forced.
"""
import shutil
import sys
from pathlib import Path

from jira2dashboard.paths import OUT_DIR

TOOL_CACHES = (".coverage", ".pytest_cache")


def is_kept(entry: Path) -> bool:
    return entry.name in TOOL_CACHES or (entry.is_file() and entry.suffix.lower() == ".xlsm")


def intermediates(out_dir: Path) -> list[Path]:
    if not out_dir.is_dir():
        return []
    return sorted(entry for entry in out_dir.iterdir() if not is_kept(entry))


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


def main(out_dir: Path = OUT_DIR) -> int:
    removed, in_use = clean(out_dir)
    print(f"clean: {len(removed)} intermediate(s) removed from {out_dir}")
    for entry in in_use:
        print(f"clean: kept {entry.name} (in use)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
