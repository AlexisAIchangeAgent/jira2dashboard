"""testbench.keep_log: the logs of each Excel test are kept in a file of their own, in the folder logs of the test workbooks."""
from jira2dashboard.testbench import keep_log as K


def write_log(folder, name, text):
    folder.mkdir(exist_ok=True)
    (folder / name).write_bytes(text.encode("ascii"))


def test_each_log_of_the_folder_gives_the_lines_of_the_test_and_not_the_ones_before(tmp_path):
    write_log(tmp_path, K.LOG_NAME, "RUN earlier\r\n")
    marks = K.mark(tmp_path)
    write_log(tmp_path, K.LOG_NAME, "RUN earlier\r\nRUN this test\r\nERROR J2D-CFG-002 | x\r\n")
    write_log(tmp_path, "vba-tests.log", "TEST T_Log\r\n")                     # a log that was not there before
    kept = K.keep(tmp_path, marks, "test_a[one]", "# test_a | passed | 1.000 s")
    assert kept == tmp_path / "excel-test_test_a_one.log"
    assert kept.read_bytes().decode("utf-8") == (
        "# test_a | passed | 1.000 s\r\n"
        "--- jira2dashboard.log: the lines of this test\r\nRUN this test\r\nERROR J2D-CFG-002 | x\r\n"
        "--- vba-tests.log: the lines of this test\r\nTEST T_Log\r\n")


def test_the_logs_of_the_temporary_folder_of_the_test_are_added_whole(tmp_path):
    other = tmp_path / "tmp" / "python.log"
    other.parent.mkdir()
    other.write_text("RUN extract\nRUN end ok\n", encoding="utf-8")
    kept = K.keep(tmp_path / "logs", {}, "t", "# t", [other, tmp_path / "gone.log"])
    assert "--- python.log: written in the temporary folder of the test\r\nRUN extract\r\nRUN end ok\r\n" in \
        kept.read_bytes().decode("utf-8")


def test_a_log_that_was_rotated_is_kept_whole(tmp_path):
    write_log(tmp_path, K.LOG_NAME, "RUN new file\n")
    assert "RUN new file" in K.keep(tmp_path, {K.LOG_NAME: 5000}, "t", "# t").read_text(encoding="utf-8")


def test_a_test_without_any_log_line_still_has_its_file(tmp_path):
    assert K.mark(tmp_path / "none") == {}
    write_log(tmp_path, K.LOG_NAME, "RUN a\n")
    kept = K.keep(tmp_path, K.mark(tmp_path), "t", "# t | passed")
    assert kept.read_text(encoding="utf-8") == "# t | passed\n--- no log line during this test\n"


def test_the_kept_files_are_not_logs_to_split(tmp_path):
    write_log(tmp_path, "excel-test_old.log", "an older kept file\n")
    assert K.mark(tmp_path) == {} and K.added(tmp_path, {}) == []


def test_a_file_that_cannot_be_written_is_not_an_error(tmp_path):
    (tmp_path / K.file_name("t")).mkdir()                                     # a folder where the file should be
    assert K.keep(tmp_path, {}, "t", "# t") is None


def test_the_file_name_is_safe_for_windows():
    assert K.file_name("test_x[a/b: c]") == "excel-test_test_x_a_b_c.log"
    assert len(K.file_name("x" * 500)) <= 125
