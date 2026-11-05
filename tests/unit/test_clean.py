"""Cleaning the folder out before a build: the workbooks stay, the intermediate files go."""
from jira2dashboard import clean


def fill(out):
    (out / "jira2dashboard.xlsm").write_bytes(b"x")
    (out / "vba_tests_v4.XLSM").write_bytes(b"x")
    (out / "~$jira2dashboard.xlsm").write_bytes(b"lock")      # Excel's lock file of an open workbook
    (out / ".coverage").write_bytes(b"x")
    (out / ".pytest_cache").mkdir()
    (out / "excel10_1.log").write_text("log")
    (out / "modTests.bak").write_text("old")
    (out / "tmp_export").mkdir()
    (out / "tmp_export" / "done.flag").write_text("done")
    (out / "release").mkdir()
    (out / "release" / "jira2dashboard.xlsm").write_bytes(b"x")   # a workbook below out is an intermediate copy


def names(paths):
    return sorted(p.name for p in paths)


def test_intermediates_are_everything_but_the_workbooks_the_lock_files_and_the_tool_caches(tmp_path):
    fill(tmp_path)
    assert names(clean.intermediates(tmp_path)) == ["excel10_1.log", "modTests.bak", "release", "tmp_export"]


def test_clean_removes_the_intermediates_and_keeps_the_rest(tmp_path):
    fill(tmp_path)
    removed, in_use = clean.clean(tmp_path)
    assert names(removed) == ["excel10_1.log", "modTests.bak", "release", "tmp_export"]
    assert in_use == []
    assert sorted(p.name for p in tmp_path.iterdir()) == [
        ".coverage", ".pytest_cache", "jira2dashboard.xlsm", "vba_tests_v4.XLSM", "~$jira2dashboard.xlsm"]


def test_a_file_still_open_is_kept_and_reported(tmp_path):
    log = tmp_path / "running.log"
    log.write_text("being written")
    with open(log, "rb"):                                   # Windows refuses to delete an open file
        removed, in_use = clean.clean(tmp_path)
    assert (removed, names(in_use)) == ([], ["running.log"])
    assert log.exists()


def test_a_missing_folder_has_nothing_to_clean(tmp_path):
    assert clean.clean(tmp_path / "out") == ([], [])


def test_main_reports_what_it_did(tmp_path, capsys):
    fill(tmp_path)
    assert clean.main(tmp_path) == 0
    assert capsys.readouterr().out == f"clean: 4 intermediate(s) removed from {tmp_path}\n"


def test_main_names_the_files_it_could_not_remove(tmp_path, capsys):
    log = tmp_path / "running.log"
    log.write_text("being written")
    with open(log, "rb"):
        clean.main(tmp_path)
    assert "clean: kept running.log (in use)" in capsys.readouterr().out
