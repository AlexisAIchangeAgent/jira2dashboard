"""scripts\\start-mock.ps1 and scripts\\stop-mock.ps1: start the fake Jira in the background, kill only what was started.

No Excel: a fake Jira of 30 issues with no window (-Hidden), on a free port of its own for each test.
"""
import os
import socket
import subprocess
import sys
import time
from pathlib import Path

import pytest

from jira2dashboard.paths import REPO_ROOT

SCRIPTS = REPO_ROOT / "scripts"


def free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def pid_file(port: int) -> Path:
    return Path(os.environ["TEMP"]) / "jira2dashboard" / f"mock-jira-{port}.pid"


def script(name: str, *args) -> subprocess.CompletedProcess:
    return subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(SCRIPTS / name),
                           *map(str, args)], capture_output=True, text=True, timeout=120)


def start(port: int) -> subprocess.CompletedProcess:
    return script("start-mock.ps1", "-Port", port, "-Issues", 30, "-Hidden")


def stop(port: int) -> subprocess.CompletedProcess:
    return script("stop-mock.ps1", "-Port", port)


def answers(port: int) -> bool:
    try:
        with socket.create_connection(("127.0.0.1", port), timeout=0.5):
            return True
    except OSError:
        return False


def closes(port: int, seconds: float = 5) -> bool:
    deadline = time.monotonic() + seconds
    while answers(port):
        if time.monotonic() > deadline:
            return False
        time.sleep(0.1)
    return True


@pytest.fixture
def port():
    port = free_port()
    yield port
    stop(port)                                          # never leave a fake Jira behind, whatever the test did


def test_start_then_stop(port):
    started = start(port)
    assert started.returncode == 0, started.stdout + started.stderr
    assert answers(port) and pid_file(port).exists()
    stopped = stop(port)
    assert stopped.returncode == 0 and "stopped" in stopped.stdout
    assert closes(port) and not pid_file(port).exists()
    again = stop(port)
    assert again.returncode == 0 and "no fake Jira started by these scripts" in again.stdout


def test_a_port_held_by_another_program_is_refused_and_left_alone(port):
    with socket.socket() as other:
        other.bind(("127.0.0.1", port))
        other.listen()
        started = start(port)
        assert started.returncode == 1 and "held by another program" in started.stdout
        assert answers(port) and not pid_file(port).exists()


def test_stop_never_kills_a_process_that_reused_the_pid(port):
    """A pid file whose start time is not the one of the process: that process is not the fake Jira, it stays."""
    other = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(60)"])
    try:
        pid_file(port).parent.mkdir(parents=True, exist_ok=True)
        pid_file(port).write_text(f"{other.pid} 1", encoding="ascii")
        stopped = stop(port)
        assert stopped.returncode == 0 and "had already stopped" in stopped.stdout
        assert other.poll() is None and not pid_file(port).exists()
    finally:
        other.kill()
        other.wait()


@pytest.mark.parametrize("path", sorted(SCRIPTS.glob("*.ps1")), ids=lambda p: p.name)
def test_the_scripts_are_ascii(path):
    """PowerShell 5.1 reads a file without BOM as ANSI: any other character would be read wrong."""
    path.read_bytes().decode("ascii")
