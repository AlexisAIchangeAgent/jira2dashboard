"""testbench.keep_log: the VBA log of an Excel scenario is kept in its own file, in the logs folder of the workbook."""
from jira2dashboard.testbench import keep_log as K


def write_log(folder, text):
    (folder / "logs").mkdir(exist_ok=True)
    (folder / "logs" / K.LOG_NAME).write_text(text, encoding="ascii")


def test_the_lines_of_the_scenario_are_kept_next_to_the_log_of_the_workbook_and_the_ones_before_are_not(tmp_path):
    write_log(tmp_path, "RUN earlier scenario\n")
    start = K.mark(tmp_path)
    write_log(tmp_path, "RUN earlier scenario\nRUN this scenario\nERROR J2D-CFG-002 | x\n")
    kept = K.keep(tmp_path, start, "test_a[one]")
    assert kept == tmp_path / "logs" / "excel-test_test_a_one.log"           # relative to the workbook: <folder>\logs
    assert kept.read_text(encoding="ascii") == "RUN this scenario\nERROR J2D-CFG-002 | x\n"


def test_a_log_that_was_rotated_is_kept_whole(tmp_path):
    write_log(tmp_path, "RUN new file\n")
    assert K.keep(tmp_path, 5000, "t").read_text(encoding="ascii") == "RUN new file\n"


def test_nothing_is_kept_when_there_is_no_log_or_no_new_line(tmp_path):
    assert K.mark(tmp_path) == 0
    assert K.keep(tmp_path, 0, "t") is None
    write_log(tmp_path, "RUN a\n")
    assert K.keep(tmp_path, K.mark(tmp_path), "t") is None


def test_a_file_that_cannot_be_written_is_not_an_error(tmp_path):
    write_log(tmp_path, "RUN a\n")
    (tmp_path / "logs" / K.file_name("t")).mkdir()                            # a folder where the file should be
    assert K.keep(tmp_path, 0, "t") is None


def test_the_file_name_is_safe_for_windows():
    assert K.file_name("test_x[a/b: c]") == "excel-test_test_x_a_b_c.log"
    assert len(K.file_name("x" * 500)) <= 125
