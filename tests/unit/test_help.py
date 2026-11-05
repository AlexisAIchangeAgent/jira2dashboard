"""Every script answers --help: how it works and its arguments, before any clean, Excel or network.

Python entry points are parsed in process (fast); the PowerShell scripts run in a powershell.exe each.
"""
import argparse
import os
import re
import runpy
import subprocess
import sys

import pytest

from jira2dashboard import build_workbook, clean, mock_workbook
from jira2dashboard.cli import args, check_machine
from jira2dashboard.mockjira import __main__ as mock_jira
from jira2dashboard.paths import CHILDREN_DIR, REPO_ROOT
from jira2dashboard.standalone import __main__ as extract
from jira2dashboard.testbench import run_vba_tests

SCRIPTS = REPO_ROOT / "scripts"
PYTHON_COMMANDS = {                      # run.ps1 command -> (parse function, a word of its description)
    "build": (build_workbook.parse, "out\\release\\jira2dashboard.xlsm"),
    "vba-tests": (run_vba_tests.parse, "--bisect"),
    "clean": (clean.parse, "out\\inter"),
    "check-machine": (check_machine.parse, "Excel"),
    "mock-jira": (mock_jira.parse, "--anomalies"),
    "mock-workbook": (mock_workbook.parse, "jira2dashboard_mock.xlsm"),
    "extract": (extract.parse, "jira2dashboard.xlsx"),
}


def help_of(parse, capsys) -> str:
    with pytest.raises(SystemExit) as stop:
        parse(["--help"])
    assert stop.value.code == 0
    return capsys.readouterr().out


@pytest.mark.parametrize("command", PYTHON_COMMANDS)
def test_every_python_command_explains_itself_and_its_arguments(command, capsys):
    parse, word = PYTHON_COMMANDS[command]
    text = help_of(parse, capsys)
    assert text.startswith(f"usage: scripts\\run.ps1 {command}")
    assert word in text and "-h, --help" in text


@pytest.mark.parametrize("child", sorted(CHILDREN_DIR.glob("_*.py")), ids=lambda p: p.name)
def test_every_child_says_it_is_internal(child, capsys, monkeypatch):
    monkeypatch.setattr(sys, "argv", [str(child), "--help"])
    with pytest.raises(SystemExit) as stop:
        runpy.run_path(str(child), run_name="__main__")
    assert stop.value.code == 0
    assert "Internal" in capsys.readouterr().out


def test_a_wrong_argument_is_refused_before_anything_runs(capsys):
    with pytest.raises(SystemExit) as stop:
        build_workbook.parse(["--out", "..\\elsewhere.xlsm"])
    assert stop.value.code == 2
    assert "a bare file name ending in .xlsm" in capsys.readouterr().err


def test_out_takes_a_bare_xlsm_name():
    assert args.bare_xlsm("try.xlsm") == "try.xlsm"
    for wrong in ("try.xlsx", "sub\\try.xlsm", "C:\\try.xlsm"):
        with pytest.raises(argparse.ArgumentTypeError):
            args.bare_xlsm(wrong)


MODULES = ("build_workbook", "testbench.run_vba_tests", "clean", "cli.check_machine", "mockjira", "mock_workbook",
           "standalone")


@pytest.mark.parametrize("module", MODULES)
def test_help_from_the_command_line_runs_nothing(module):
    """--help is parsed before any work: no clean of out\\inter, no Excel, no server that never returns."""
    env = os.environ | {"PYTHONPATH": str(REPO_ROOT / "src")}
    done = subprocess.run([sys.executable, "-m", f"jira2dashboard.{module}", "--help"], capture_output=True,
                          text=True, env=env, timeout=60)
    assert done.returncode == 0 and done.stdout.startswith("usage: scripts\\run.ps1 ")
    assert "clean:" not in done.stdout and "built" not in done.stdout and "Fake Jira:" not in done.stdout


def powershell(script, *arguments) -> subprocess.CompletedProcess:
    return subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(SCRIPTS / script),
                           *arguments], capture_output=True, encoding="oem", errors="replace", timeout=120)


def synopsis(path) -> str:
    text = path.read_text(encoding="ascii")
    return re.search(r"\.SYNOPSIS\n(.+)\n", text).group(1)


@pytest.mark.parametrize("script", sorted(p.name for p in SCRIPTS.glob("*.ps1")))
def test_every_powershell_script_explains_itself(script):
    done = powershell(script, "--help")
    assert done.returncode == 0, done.stdout + done.stderr
    if script == "run.ps1":
        assert "Usage: .\\scripts\\run.ps1 <command>" in done.stdout
    else:
        assert synopsis(SCRIPTS / script)[:60] in " ".join(done.stdout.split())


def test_run_lists_every_command_it_runs():
    text = (SCRIPTS / "run.ps1").read_text(encoding="ascii")
    commands = re.findall(r'^\s+"([a-z-]+)"\s+\{ & \$python', text, flags=re.MULTILINE)
    listed = powershell("run.ps1", "--help").stdout
    assert commands and all(re.search(rf"^\s+{c}\s", listed, flags=re.MULTILINE) for c in commands)
    assert set(PYTHON_COMMANDS) <= set(commands)


def test_a_powershell_script_refuses_an_unknown_argument_and_takes_minus_help():
    assert powershell("stop-mock.ps1", "-Help").returncode == 0
    wrong = powershell("stop-mock.ps1", "--bogus")
    assert wrong.returncode == 2 and "--help" in wrong.stdout
