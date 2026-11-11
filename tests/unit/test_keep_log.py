"""testbench.keep_log: the VBA log of an Excel scenario is kept for after the run."""
from jira2dashboard.testbench import keep_log as K


def write_log(folder, text):
    (folder / "logs").mkdir(exist_ok=True)
    (folder / "logs" / K.LOG_NAME).write_text(text, encoding="ascii")


def test_the_lines_of_the_scenario_are_kept_and_the_ones_before_are_not(tmp_path):
    book, local = tmp_path / "book", tmp_path / "local"
    book.mkdir()
    write_log(book, "RUN earlier scenario\n")
    start = K.mark(book)
    write_log(book, "RUN earlier scenario\nRUN this scenario\nERROR J2D-CFG-002 | x\n")
    kept = K.keep(book, start, "test_a[one]", str(local))
    assert kept == local / "jira2dashboard" / "logs" / "excel-tests" / "test_a_one.log"
    assert kept.read_text(encoding="ascii") == "RUN this scenario\nERROR J2D-CFG-002 | x\n"


def test_a_log_that_was_rotated_is_kept_whole(tmp_path):
    write_log(tmp_path, "RUN new file\n")
    kept = K.keep(tmp_path, 5000, "t", str(tmp_path / "local"))
    assert kept.read_text(encoding="ascii") == "RUN new file\n"


def test_nothing_is_kept_when_there_is_no_log_or_no_new_line(tmp_path):
    assert K.mark(tmp_path) == 0
    assert K.keep(tmp_path, 0, "t", str(tmp_path / "local")) is None
    write_log(tmp_path, "RUN a\n")
    assert K.keep(tmp_path, K.mark(tmp_path), "t", str(tmp_path / "local")) is None


def test_a_folder_that_cannot_be_written_is_not_an_error(tmp_path):
    write_log(tmp_path, "RUN a\n")
    blocker = tmp_path / "blocker"
    blocker.write_text("a file where the folder should be")
    assert K.keep(tmp_path, 0, "t", str(blocker)) is None


def test_the_file_name_is_safe_for_windows():
    assert K.file_name("test_x[a/b: c]") == "test_x_a_b_c.log"
    assert len(K.file_name("x" * 500)) <= 124
