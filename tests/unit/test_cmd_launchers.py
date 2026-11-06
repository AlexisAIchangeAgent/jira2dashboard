"""The double-click launchers scripts\\<name>.cmd: one per PowerShell script, run with the execution policy bypassed.

A double-click must show the output until a key is pressed (the window would close at once otherwise); from a console or a
script there is no wait. --help explains the launcher and then the script. Nothing here starts Excel or a server.
"""
import os
import re
import subprocess

import pytest

from jira2dashboard.paths import REPO_ROOT

SCRIPTS = REPO_ROOT / "scripts"
NAMES = sorted(p.stem for p in SCRIPTS.glob("*.ps1"))
WAIT_MESSAGE = "Press any key to close this window"


def launcher(name: str):
    return SCRIPTS / f"{name}.cmd"


def run(command, *, env=None, cwd=None, stdin_text="\r\n") -> subprocess.CompletedProcess:
    """`command` is a list, or a raw command line (what Explorer gives cmd.exe on a double-click)."""
    return subprocess.run(command, capture_output=True, encoding="oem", errors="replace", timeout=120, cwd=cwd,
                          env={**os.environ, **(env or {})}, input=stdin_text)


def from_console(name, *arguments, env=None):
    """As a person types it in a console, or a script calls it by its short name: no double-click, no wait."""
    return run(["cmd.exe", "/c", f".\{name}.cmd", *arguments], cwd=SCRIPTS, env=env)


def double_click(name):
    """The command line Explorer gives cmd.exe: C:\WINDOWS\system32\cmd.exe /c ""<path>" " (no arguments)."""
    return run(f'{os.environ["COMSPEC"]} /c ""{launcher(name)}" "')


def test_every_powershell_script_has_a_launcher_and_there_is_nothing_else():
    cmd_names = {p.stem for p in SCRIPTS.glob("*.cmd")}
    assert set(NAMES) <= cmd_names
    assert cmd_names - set(NAMES) == {"_launch"}


@pytest.mark.parametrize("path", sorted(SCRIPTS.glob("*.cmd")), ids=lambda p: p.name)
def test_the_launchers_are_ascii_with_windows_line_endings(path):
    data = path.read_bytes()
    data.decode("ascii")
    assert data.startswith(b"@echo off") and b"\n" in data
    assert not re.search(rb"(?<!\r)\n", data), "a bare LF can break goto and call labels"


@pytest.mark.parametrize("name", NAMES)
def test_help_explains_the_launcher_then_the_script(name):
    done = from_console(name, "--help")
    text = " ".join(done.stdout.split())
    assert done.returncode == 0, done.stdout + done.stderr
    assert f"{name}.cmd - launcher of scripts\\{name}.ps1" in text
    what = re.search(r'set "J2D_WHAT=(.+)"', launcher(name).read_text(encoding="ascii")).group(1)
    assert not re.search(r"[<>|&^%!]", what), "a character cmd would take for an operator"
    assert what in text, "the one-line description must be printed"
    assert "Double-click" in text and "ExecutionPolicy Bypass" in text and "J2D_PAUSE" in text
    assert "Exit code" not in done.stdout and WAIT_MESSAGE not in done.stdout          # a console does not wait
    if name == "run":
        assert "Usage: .\\scripts\\run.ps1 <command>" in text
    else:
        ps1 = (SCRIPTS / f"{name}.ps1").read_text(encoding="ascii")
        assert re.search(r"\.SYNOPSIS\n(.+)\n", ps1).group(1)[:60] in text


@pytest.mark.parametrize("flag", ["-h", "-help", "/?"])
def test_the_other_spellings_of_help_work(flag):
    done = from_console("stop-mock", flag)
    assert done.returncode == 0 and "stop-mock.cmd - launcher of" in done.stdout


def test_a_double_click_waits_for_a_key_and_says_the_exit_code():
    done = double_click("release")                      # no version: it only explains itself, it never builds a release
    assert WAIT_MESSAGE in done.stdout and "Exit code: 2" in done.stdout
    assert done.returncode == 2


def test_the_wait_can_be_forced_from_a_console():
    done = from_console("stop-mock", "--help", env={"J2D_PAUSE": "1"})
    assert WAIT_MESSAGE in done.stdout and "Exit code: 0" in done.stdout


def test_release_without_arguments_is_not_run_and_says_how_to_run_it():
    done = from_console("release")
    assert done.returncode == 2
    assert "Not run" in done.stdout and "release.cmd -Version v1.0.0" in done.stdout
    assert "clean:" not in done.stdout and "built" not in done.stdout


def test_the_exit_code_of_the_script_comes_back():
    done = from_console("run", "no-such-command")
    assert done.returncode != 0 and "no-such-command" in done.stdout + done.stderr


def test_the_arguments_reach_the_script():
    wrong = from_console("stop-mock", "--bogus")
    assert wrong.returncode == 2 and "--help" in wrong.stdout


def test_every_powershell_call_of_the_helper_bypasses_the_policy_for_that_run_only():
    calls = [line for line in (SCRIPTS / "_launch.cmd").read_text(encoding="ascii").splitlines()
             if line.startswith("powershell ")]
    assert len(calls) == 3                                       # the script, the console of activate, the help
    assert all("-NoProfile -ExecutionPolicy Bypass" in line for line in calls)
