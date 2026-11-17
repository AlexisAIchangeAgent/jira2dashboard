"""The PNG of the charts that a test exports to be looked at (AL, 2026-11-17): they stay after `build` / `vba-tests`
(`paths.CHARTS_DIR`, outside `out\\inter`) and each run replaces them: one sub-folder per scenario, emptied of its
old PNG before the export, and a note `_exported.txt` written AFTER it, so that a run that failed before the export
shows an old time (or none) instead of images that look current."""
import datetime as dt
import subprocess
from pathlib import Path

NOTE_NAME = "_exported.txt"


def prepare(folder: Path) -> None:
    """Make the folder and remove the PNG and the note of the previous run (a chart renamed or removed leaves nothing)."""
    folder.mkdir(parents=True, exist_ok=True)
    for old in list(folder.glob("*.png")) + [folder / NOTE_NAME]:
        old.unlink(missing_ok=True)


def commit_of(repo: Path) -> str:
    """The short hash of HEAD, or "unknown" (no git, not a repository)."""
    try:
        done = subprocess.run(["git", "rev-parse", "--short", "HEAD"], cwd=repo, capture_output=True, text=True,
                              timeout=10, check=True)
        return done.stdout.strip() or "unknown"
    except (OSError, subprocess.SubprocessError):
        return "unknown"


def note_text(scenario: str, names: list[str], commit: str, now: dt.datetime) -> str:
    return (f"charts of the scenario {scenario}\n"
            f"exported (UTC): {now:%Y-%m-%d %H:%M:%S}\n"
            f"commit: {commit}\n"
            f"images: {len(names)}\n" + "".join(f"  {n}\n" for n in sorted(names)))


def write_note(folder: Path, scenario: str, names: list[str], commit: str) -> Path:
    note = folder / NOTE_NAME
    note.write_text(note_text(scenario, names, commit, dt.datetime.now(dt.timezone.utc)), encoding="utf-8")
    return note
