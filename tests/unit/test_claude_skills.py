"""The script of the skill read-logs (.claude/skills/read-logs/summarize_logs.py): what it keeps of each kind of log.
Skipped where .claude is missing (the release zip)."""
import importlib.util

import pytest

from jira2dashboard.paths import REPO_ROOT

SCRIPT = REPO_ROOT / ".claude" / "skills" / "read-logs" / "summarize_logs.py"


@pytest.fixture(scope="module")
def S():
    if not SCRIPT.is_file():
        pytest.skip(".claude/skills is not in this copy of the code")
    spec = importlib.util.spec_from_file_location("summarize_logs", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


VBA = """2026-11-11 20:24:18.906 RUN button Extract from Jira and compute | Excel 16.0 | wb.xlsm | jira2dashboard commit 4df0f08
2026-11-11 20:24:18.921 PARAM BaseUrl = (set)
2026-11-11 20:24:19.617 END search page 1: 100 issues | 0.336 s
2026-11-11 20:24:21.617 END flow sheets: formulas ItemFlow | 1.938 s
2026-11-11 20:24:22.000 RETRY HTTP 429 at attempt 1 of 5: next attempt in 2.345 s
2026-11-11 20:24:35.664 RUN end ok | 16.805 s
2026-11-11 20:25:14.968 RUN button Read the Jira configuration | Excel 16.0 | wb.xlsm | jira2dashboard commit 4df0f08
2026-11-11 20:25:15.000 ERROR J2D-JIRA-401 | BtnReadJiraConfiguration | vbObjectError+602 | (description not logged)
2026-11-11 20:25:15.010 RUN end stopped J2D-JIRA-401 | 0.042 s
2026-11-11 20:26:00.000 RUN button Extract from Jira and compute | Excel 16.0 | wb.xlsm | jira2dashboard commit 4df0f08
2026-11-11 20:26:05.000 BEGIN search page 30
"""


def test_the_vba_log_gives_runs_errors_retries_slow_steps_and_a_run_that_did_not_end(S):
    report = S.summarize("jira2dashboard.log", VBA.splitlines())
    text = "\n".join(report)
    assert report[0] == "== jira2dashboard.log"
    assert "RUN end stopped J2D-JIRA-401 | 0.042 s" in text and "ERROR J2D-JIRA-401" in text
    assert "RETRY HTTP 429 at attempt 1 of 5" in text
    assert "slow: 1.938 s  2026-11-11 20:24:21.617 flow sheets: formulas ItemFlow" in text
    assert "search page 1" not in text                                 # under half a second
    assert "PARAM" not in text
    assert report[-2].startswith("NOT ENDED") and "20:26:00.000" in report[-2]
    assert S.codes(report) == ["J2D-JIRA-401"]


def test_a_kept_excel_test_log_names_a_failed_test_and_a_passed_one_says_nothing(S):
    failed = S.summarize("x.log", ["# tests/steps/t.py::test_a | failed | 45.191 s | 2026-11-11 19:25:00"])
    assert failed == ["== x.log", "TEST FAILED: tests/steps/t.py::test_a (45.191 s | 2026-11-11 19:25:00)"]
    assert S.summarize("y.log", ["# tests/steps/t.py::test_b | passed | 1.000 s | 2026-11-11 19:25:00",
                                 "--- no log line during this test"]) == []


def test_python_log_and_pytest_lines_are_read(S):
    report = S.summarize("run.log", [
        "2026-11-11 18:38:31,228 INFO RUN extract | Python 3.14.7 | wb.xlsm | jira2dashboard unknown",
        "2026-11-11 18:38:31,300 WARNING search: GET /rest/api/3/search/jql -> HTTP 429 (attempt 1 of 5)",
        "2026-11-11 18:38:32,239 INFO RUN end ok: C:\\out\\jira2dashboard.xlsx | 1.031 s",
        "FAILED tests/steps/test_build_workbook.py::test_x - RuntimeError",
        "===== 1 failed, 43 passed, 969 deselected in 301.38s (0:05:01) ====="])
    text = "\n".join(report)
    assert "WARNING search: GET" in text and "RUN end ok" in text and "NOT ENDED" not in text
    assert "pytest: FAILED tests/steps/test_build_workbook.py::test_x - RuntimeError" in text
    assert "pytest: ===== 1 failed, 43 passed" in text


def test_folders_and_files_are_read_and_nothing_found_is_said(S, tmp_path, capsys):
    (tmp_path / "a.log").write_text("2026-11-11 20:24:18.906 RUN end ok | 1.000 s\n", encoding="ascii")
    (tmp_path / "b.txt").write_text("not a log", encoding="ascii")
    assert S.files([tmp_path, tmp_path / "missing.log"]) == [tmp_path / "a.log"]
    S.main([str(tmp_path / "empty")])
    assert "Nothing to report" in capsys.readouterr().out


# --- the skill vba-mutation: .claude/skills/vba-mutation/mutate.py ---------------------------------------------------------------

MUTATE = REPO_ROOT / ".claude" / "skills" / "vba-mutation" / "mutate.py"


@pytest.fixture
def mutate(tmp_path, monkeypatch):
    if not MUTATE.is_file():
        pytest.skip(".claude/skills is not in this copy of the code")
    monkeypatch.setenv("J2D_MUTATION_DIR", str(tmp_path / "keep"))               # never the real folder of %TEMP%
    spec = importlib.util.spec_from_file_location("mutate", MUTATE)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@pytest.fixture
def bas(tmp_path):
    path = tmp_path / "modX.bas"
    path.write_bytes(b"Option Explicit\nPublic Function F() As Long\n    F = 1\nEnd Function\n")
    return path


def test_a_mutation_is_applied_once_and_restored_byte_for_byte(mutate, bas, capsys):
    good = bas.read_bytes()
    assert mutate.main(["apply", str(bas), "--old", "F = 1", "--new", "F = 2"]) == 0
    assert bas.read_bytes() == good.replace(b"F = 1", b"F = 2") and b"\r" not in bas.read_bytes()       # LF stays LF
    assert mutate.main(["status"]) == 1 and "MUTATED: modX.bas" in capsys.readouterr().out
    assert mutate.main(["restore", str(bas)]) == 0 and bas.read_bytes() == good
    assert mutate.main(["status"]) == 0


def test_a_mutation_that_does_not_apply_exactly_once_changes_nothing(mutate, bas):
    good = bas.read_bytes()
    for old in ("not there", "F"):                                                # 0 times, 3 times
        with pytest.raises(SystemExit, match="exactly once"):
            mutate.main(["apply", str(bas), "--old", old, "--new", "x"])
    with pytest.raises(SystemExit, match="no mutation"):
        mutate.main(["apply", str(bas), "--old", "F = 1", "--new", "F = 1"])
    with pytest.raises(SystemExit, match="ASCII"):                                # a VBA module is ASCII
        mutate.main(["apply", str(bas), "--old", "F = 1", "--new", "F = é"])
    assert bas.read_bytes() == good and mutate.main(["status"]) == 0


def test_a_module_already_mutated_is_not_mutated_again_and_restore_without_a_mutation_is_harmless(mutate, bas, capsys):
    good = bas.read_bytes()
    assert mutate.main(["restore", str(bas)]) == 0 and "nothing to restore" in capsys.readouterr().out
    mutate.main(["apply", str(bas), "--old", "F = 1", "--new", "F = 2"])
    with pytest.raises(SystemExit, match="already mutated"):
        mutate.main(["apply", str(bas), "--old", "F = 2", "--new", "F = 3"])     # the good text must not be overwritten
    mutate.main(["restore", str(bas)])
    assert bas.read_bytes() == good


def test_only_a_bas_module_is_touched(mutate, tmp_path):
    other = tmp_path / "notes.txt"
    other.write_text("a")
    with pytest.raises(SystemExit, match="not a .bas module"):
        mutate.main(["apply", str(other), "--old", "a", "--new", "b"])
