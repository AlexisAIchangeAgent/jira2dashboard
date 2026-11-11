"""Cleaning the folder out before a build: out\\release and out\\manual_test stay, out\\inter is emptied, nothing else
stays at the root."""
from jira2dashboard import clean


def fill(out):
    release, inter, manual = out / "release", out / "inter", out / "manual_test"
    (release / "v1.0.0").mkdir(parents=True)
    (release / "jira2dashboard.xlsm").write_bytes(b"x")                  # the workbook to deliver
    (release / "v1.0.0" / "jira2dashboard-v1.0.0.zip").write_bytes(b"x")  # a release package
    manual.mkdir()
    (manual / "jira2dashboard4manualtest.xlsm").write_bytes(b"x")        # the copy of manual-test.ps1
    inter.mkdir()
    (inter / "vba_tests.xlsm").write_bytes(b"x")
    (inter / "~$vba_tests.xlsm").write_bytes(b"lock")                    # Excel's lock file of an open workbook
    (inter / ".coverage").write_bytes(b"x")
    (inter / ".pytest_cache").mkdir()
    (inter / "dashboard").mkdir()
    (inter / "dashboard" / "chart.png").write_bytes(b"x")
    (inter / "logs").mkdir()
    (inter / "logs" / "jira2dashboard.log").write_text("RUN")              # the logs of the test workbooks stay
    (out / "jira2dashboard.xlsm").write_bytes(b"x")                      # leftovers of the old layout
    (out / "excel10_1.log").write_text("log")


def names(out, paths):
    return sorted(p.relative_to(out).as_posix() for p in paths)


def tree(out):
    return sorted(p.relative_to(out).as_posix() for p in out.rglob("*"))


def test_intermediates_are_the_content_of_inter_but_the_tool_caches_and_everything_else_at_the_root(tmp_path):
    fill(tmp_path)
    assert names(tmp_path, clean.intermediates(tmp_path)) == [
        "excel10_1.log", "inter/dashboard", "inter/vba_tests.xlsm", "inter/~$vba_tests.xlsm",
        "jira2dashboard.xlsm"]


def test_clean_keeps_release_the_manual_test_copy_the_tool_caches_and_the_logs(tmp_path):
    fill(tmp_path)
    removed, in_use = clean.clean(tmp_path)
    assert len(removed) == 5 and in_use == []
    assert tree(tmp_path) == [
        "inter", "inter/.coverage", "inter/.pytest_cache", "inter/logs", "inter/logs/jira2dashboard.log",
        "manual_test", "manual_test/jira2dashboard4manualtest.xlsm",
        "release", "release/jira2dashboard.xlsm", "release/v1.0.0", "release/v1.0.0/jira2dashboard-v1.0.0.zip"]


def test_a_file_still_open_is_kept_and_reported(tmp_path):
    (tmp_path / "inter").mkdir()
    log = tmp_path / "inter" / "running.log"
    log.write_text("being written")
    with open(log, "rb"):                                   # Windows refuses to delete an open file
        removed, in_use = clean.clean(tmp_path)
    assert (removed, names(tmp_path, in_use)) == ([], ["inter/running.log"])
    assert log.exists()


def test_a_missing_folder_has_nothing_to_clean(tmp_path):
    assert clean.clean(tmp_path / "out") == ([], [])


def test_main_reports_what_it_did(tmp_path, capsys):
    fill(tmp_path)
    assert clean.main(tmp_path) == 0
    assert capsys.readouterr().out == f"clean: 5 intermediate(s) removed from {tmp_path}\n"


def test_main_names_the_files_it_could_not_remove(tmp_path, capsys):
    log = tmp_path / "running.log"
    log.write_text("being written")
    with open(log, "rb"):
        clean.main(tmp_path)
    assert "clean: kept running.log (in use)" in capsys.readouterr().out
