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
    for name in ("old_chart.png", "old_chart.svg", CI.INDEX_NAME, CI.NOTE_NAME, "mine.txt"):
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


GROUPS = [("Charts", [("Velocity <per iteration>", "velocity"), ("Flow Time", "flow_time")]), ("IterationCharts", [("Engaged", "features_engaged")])]
NOW = dt.datetime(2026, 11, 17, 9, 5, 7)


def test_the_page_shows_every_chart_as_svg_with_its_png_and_a_zoom():
    page = CI.index_html("art_metrics", GROUPS, "abc1234", NOW)
    for name in ("velocity", "flow_time", "features_engaged"):
        assert f'src="{name}.svg"' in page and f'href="{name}.png"' in page and f'href="{name}.svg"' in page
    assert page.count('class="pic" data-zoom') == 3 and 'id="zoom"' in page and "wheel" in page
    assert "3 charts" in page and "2026-11-17 09:05:07" in page and "commit abc1234" in page
    assert "<h2>Charts</h2>" in page and "<h2>IterationCharts</h2>" in page


def test_the_page_escapes_what_it_shows_and_loads_nothing_from_outside():
    page = CI.index_html("<x>", GROUPS, "a&b", NOW)
    assert "Velocity &lt;per iteration&gt;" in page and "<x>" not in page and "a&amp;b" in page
    assert "http://" not in page and "https://" not in page and "src=\"//" not in page


def test_write_index_puts_the_page_in_the_scenario_folder(tmp_path):
    page = CI.write_index(tmp_path, "flow", GROUPS, "abc1234")
    assert page == tmp_path / CI.INDEX_NAME and page.read_text(encoding="utf-8").startswith("<!doctype html>")


def test_the_root_page_links_the_scenarios_that_have_a_page(tmp_path):
    for name in ("flow", "art_metrics"):
        (tmp_path / name).mkdir()
        CI.write_index(tmp_path / name, name, GROUPS, "abc1234")
        CI.write_note(tmp_path / name, name, ["a.png"], "abc1234")
    (tmp_path / "empty").mkdir()
    page = CI.write_root_index(tmp_path).read_text(encoding="utf-8")
    assert 'href="art_metrics/index.html"' in page and 'href="flow/index.html"' in page and "empty" not in page
    assert page.index("art_metrics/") < page.index("flow/") and "exported (UTC): 20" in page


def test_the_root_page_without_a_scenario_is_not_written(tmp_path):
    assert CI.write_root_index(tmp_path) is None and not (tmp_path / CI.INDEX_NAME).exists()
