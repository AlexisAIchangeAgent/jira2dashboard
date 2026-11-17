"""The PNG of the charts that the Excel tests keep to be looked at: outside out\\inter, replaced at each run."""
import datetime as dt
import subprocess

from jira2dashboard import chart_images as CI
from jira2dashboard import clean
from jira2dashboard.paths import CHARTS_DIR, INTER_DIR, MANUAL_TEST_DIR


def test_the_images_are_kept_in_a_folder_that_clean_never_empties():
    assert CHARTS_DIR.parent == MANUAL_TEST_DIR and CHARTS_DIR.name == "charts"
    assert INTER_DIR not in CHARTS_DIR.parents


def test_clean_keeps_the_chart_images(tmp_path):
    (tmp_path / "manual_test" / "charts" / "art_metrics").mkdir(parents=True)
    kept = tmp_path / "manual_test" / "charts" / "art_metrics" / "velocity.png"
    kept.write_bytes(b"png")
    (tmp_path / "inter").mkdir()
    (tmp_path / "inter" / "x.xlsm").write_bytes(b"x")
    clean.clean(tmp_path)
    assert kept.exists() and not (tmp_path / "inter" / "x.xlsm").exists()


def test_prepare_removes_the_old_png_and_the_note_but_nothing_else(tmp_path):
    folder = tmp_path / "scenario"
    folder.mkdir()
    for name in ("old_chart.png", CI.NOTE_NAME, "mine.txt"):
        (folder / name).write_text("x")
    CI.prepare(folder)
    assert sorted(p.name for p in folder.iterdir()) == ["mine.txt"]


def test_prepare_makes_the_folder(tmp_path):
    CI.prepare(tmp_path / "a" / "b")
    assert (tmp_path / "a" / "b").is_dir()


def test_the_note_names_the_scenario_the_time_the_commit_and_the_images():
    text = CI.note_text("art_metrics", ["velocity.png", "flow_time.png"], "abc1234", dt.datetime(2026, 11, 17, 9, 5, 7))
    assert text == ("charts of the scenario art_metrics\nexported (UTC): 2026-11-17 09:05:07\ncommit: abc1234\nimages: 2\n"
                    "  flow_time.png\n  velocity.png\n")


def test_write_note_puts_the_note_next_to_the_images(tmp_path):
    note = CI.write_note(tmp_path, "flow", ["a.png"], "abc1234")
    assert note == tmp_path / CI.NOTE_NAME
    assert "commit: abc1234" in note.read_text(encoding="utf-8") and "(UTC): 20" in note.read_text(encoding="utf-8")


def test_commit_of_this_repository_is_a_short_hash_or_unknown(tmp_path):
    assert CI.commit_of(tmp_path) == "unknown" or len(CI.commit_of(tmp_path)) >= 4      # a folder that is no repository (or no git)


def test_commit_of_reads_git(monkeypatch, tmp_path):
    monkeypatch.setattr(subprocess, "run", lambda *a, **k: subprocess.CompletedProcess(a, 0, stdout="abc1234\n"))
    assert CI.commit_of(tmp_path) == "abc1234"


def test_commit_of_without_git_is_unknown(monkeypatch, tmp_path):
    def no_git(*a, **k):
        raise FileNotFoundError("git")
    monkeypatch.setattr(subprocess, "run", no_git)
    assert CI.commit_of(tmp_path) == "unknown"
