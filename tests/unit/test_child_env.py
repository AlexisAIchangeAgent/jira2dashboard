"""The Excel child processes must import jira2dashboard without `pip install` (PYTHONPATH only)."""
import ast
import os
import subprocess
import sys

import pytest

from jira2dashboard import paths
from jira2dashboard.excel import com


def test_child_env_puts_src_first_on_pythonpath():
    env = com.child_env({"PATH": "p"})
    assert env["PYTHONPATH"] == str(paths.SRC_DIR)
    assert env["PATH"] == "p"


def test_child_env_keeps_an_existing_pythonpath():
    env = com.child_env({"PYTHONPATH": "other"})
    assert env["PYTHONPATH"] == str(paths.SRC_DIR) + os.pathsep + "other"


def test_child_env_does_not_modify_the_parent_environment():
    before = os.environ.get("PYTHONPATH")
    com.child_env()
    assert os.environ.get("PYTHONPATH") == before


def test_repo_layout_constants_point_to_real_places():
    assert (paths.REPO_ROOT / "pyproject.toml").is_file()
    assert len(list(paths.VBA_DIR.glob("*.bas"))) == 18
    assert sorted(p.name for p in paths.PARAMS_DIR.glob("*.md")) == [
        "config.md", "field_mapping.md", "iterations.md", "pis.md", "projects.md", "transitions.md", "workflows.md"]
    assert sorted(p.name for p in paths.CHILDREN_DIR.glob("*.py")) == [
        "_check_child.py", "_hang_child.py", "_run_tests_child.py"]


def _jira2dashboard_imports(script):
    tree = ast.parse(script.read_text(encoding="utf-8"))
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and node.module and node.module.startswith("jira2dashboard"):
            for alias in node.names:
                yield node.module, alias.name
        elif isinstance(node, ast.Import):
            for alias in node.names:
                if alias.name.startswith("jira2dashboard"):
                    yield alias.name, None


@pytest.mark.parametrize("script", sorted(paths.CHILDREN_DIR.glob("*.py")), ids=lambda p: p.name)
def test_child_script_imports_resolve_with_child_env_only(script, tmp_path):
    imports = list(_jira2dashboard_imports(script))
    assert imports, "a child script must import the package"
    code = "; ".join(
        f"import importlib; m = importlib.import_module({mod!r}); "
        + (f"getattr(m, {name!r}, None) is not None or importlib.import_module({mod + '.' + name!r})"
           if name else "pass")
        for mod, name in imports)
    result = subprocess.run([sys.executable, "-c", code], cwd=tmp_path, env=com.child_env(),
                            capture_output=True, text=True, timeout=60)
    assert result.returncode == 0, result.stderr
