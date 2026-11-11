"""The hooks of .claude/hooks (the Claude config of the repository): the guard against stopping Excel by its name, and
the VBA guard tests after a module changed through a command. Skipped where .claude is missing (the release zip)."""
import importlib.util
import json
from pathlib import Path

import pytest

from jira2dashboard.paths import REPO_ROOT

HOOKS = REPO_ROOT / ".claude" / "hooks"


def load(name: str):
    path = HOOKS / f"{name}.py"
    if not path.is_file():
        pytest.skip(".claude/hooks is not in this copy of the code")
    spec = importlib.util.spec_from_file_location(f"hook_{name}", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@pytest.mark.parametrize("command", [
    "taskkill /IM EXCEL.EXE /F", "MSYS_NO_PATHCONV=1 taskkill //F //IM excel.exe", "Stop-Process -Name EXCEL -Force",
    "Get-Process excel | Stop-Process", "pkill -f excel",
    # a PID next to a name or a pipe does not make it safe (security review of c7be0dc)
    "taskkill /PID 123 /IM EXCEL.EXE /F", "Stop-Process -Id 1; Stop-Process -Name EXCEL", "spps -Name excel",
    "tskill excel", "wmic process where name='EXCEL.EXE' call terminate",
    "Get-CimInstance Win32_Process -Filter \"Name='EXCEL.EXE'\" | Invoke-CimMethod -MethodName Terminate",
    "(Get-Process -Id 42).Kill()  # the test Excel", "Get-Process EXCEL | % { $_.Kill() }",
    "Stop-Process -Id 5 -ProcessName EXCEL"])
def test_a_command_that_stops_excel_by_its_name_is_asked(command):
    assert load("excel_kill_guard").kills_excel_by_name(command)


@pytest.mark.parametrize("command", [
    "taskkill /PID 36044 /F", "Stop-Process -Id 1234 -Force", "tasklist | grep -c EXCEL", "taskkill /PID 99 /T",
    "powershell -File scripts/stop-mock.ps1 -Port 8099", "taskkill /PID 36044 /F  # the hung test Excel",
    "Get-Process EXCEL | Select Id, StartTime, MainWindowTitle"])
def test_a_command_that_names_one_process_or_kills_nothing_passes(command):
    assert not load("excel_kill_guard").kills_excel_by_name(command)


def test_the_guard_answers_ask_with_the_rule(monkeypatch, capsys):
    guard = load("excel_kill_guard")
    monkeypatch.setattr("sys.stdin", __import__("io").StringIO(json.dumps({"tool_input": {"command": "taskkill /IM EXCEL.EXE"}})))
    guard.main()
    answer = json.loads(capsys.readouterr().out)["hookSpecificOutput"]
    assert answer["permissionDecision"] == "ask" and "/PID" in answer["permissionDecisionReason"]


def test_a_module_changed_by_a_command_is_found_once(tmp_path):
    guard = load("bas_guard")
    vba, state = tmp_path / "vba", tmp_path / "state" / "bas_guard.json"
    vba.mkdir()
    (vba / "modA.bas").write_text("Option Explicit\n")
    assert guard.changed_modules(vba, state) == []                     # the first check only records
    assert guard.changed_modules(vba, state) == []
    (vba / "modA.bas").write_text("Option Explicit\n' changed\n")
    (vba / "modB.bas").write_text("Option Explicit\n")
    assert guard.changed_modules(vba, state) == ["modA.bas", "modB.bas"]
    assert guard.changed_modules(vba, state) == []                     # already tested
    (vba / "modB.bas").unlink()
    assert guard.changed_modules(vba, state) == ["modB.bas"]


def test_only_a_bas_of_the_vba_folder_is_an_edited_module(tmp_path):
    guard = load("bas_guard")
    vba = tmp_path / "vba"
    vba.mkdir()
    assert guard.edited_bas({"tool_input": {"file_path": str(vba / "modA.bas")}}, vba) == (vba / "modA.bas").resolve()
    assert guard.edited_bas({"tool_input": {"file_path": str(tmp_path / "modA.bas")}}, vba) is None
    assert guard.edited_bas({"tool_input": {"file_path": str(vba / "notes.md")}}, vba) is None
    assert guard.edited_bas({}, vba) is None
