"""version: the text written at build into the hidden name J2D_Version of the workbook, read back on the RUN line of the logs."""
import datetime as dt
import subprocess

from jira2dashboard import version as V

BUILT = dt.datetime(2026, 11, 11, 19, 40)


def test_the_stamp_names_the_release_the_commit_and_the_time_of_the_build():
    assert V.stamp("v1.0.0", "1a2b3c4", False, BUILT) == "v1.0.0 commit 1a2b3c4, built 2026-11-11 19:40"
    assert V.stamp(None, "1a2b3c4", True, BUILT) == "commit 1a2b3c4 with local changes, built 2026-11-11 19:40"


def test_the_name_holds_a_text_constant_with_its_quotes_doubled():
    assert V.formula('a "b"') == '="a ""b"""'


def test_the_commit_comes_from_git_and_says_when_the_tree_has_changes():
    answers = {"rev-parse": "1a2b3c4\n", "status": " M src/x.py\n"}

    def run(cmd, **_kwargs):
        return subprocess.CompletedProcess(cmd, 0, answers[cmd[3]], "")
    assert V.commit(run=run) == ("1a2b3c4", True)
    answers["status"] = ""
    assert V.commit(run=run) == ("1a2b3c4", False)


def test_without_git_or_outside_a_repository_the_commit_is_unknown():
    def missing(_cmd, **_kwargs):
        raise FileNotFoundError("git")

    def refused(cmd, **_kwargs):
        raise subprocess.CalledProcessError(128, cmd)
    assert V.commit(run=missing) == V.commit(run=refused) == ("unknown", False)


def test_the_commit_of_this_repository_is_found():
    head, _dirty = V.commit()
    assert head != V.UNKNOWN and len(head) >= 7
