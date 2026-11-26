"""A text that sends the person to a sheet names the sheet as Excel shows it (AL, 2026-11-26: "Not found in Jira: 9 project key(s) of the
sheet Projects" sent him to a sheet that is called Projects_from_user). The rule reads the string literals of the delivered code: the VBA
modules (not modTests, not the comments) and the Python package (not the docstrings, not the fake Jira, not the developer tools). A name built
with SheetName / sheet_names.of is not a literal and passes; the exceptions below are names and keys, not directions to a sheet."""
import ast
import re
from pathlib import Path

from jira2dashboard import sheet_names as S

ROOT = Path(__file__).resolve().parents[2]
NAMES = "|".join(sorted(S.USER + S.JIRA + S.EXCEL, key=len, reverse=True))
USER = "|".join(S.USER)
PATTERNS = (
    re.compile(rf"\b[Ss]heets? (?:{NAMES})\b(?!_)"),                     # "Sheet Iterations: ...", "the sheet Workflows", "the sheets JiraFields"
    re.compile(rf"\bthe (?:{NAMES}) sheet\b"),                            # "the Config sheet"
    re.compile(rf"\b(?:in|from|into|(?<!list )of) (?:{USER})\b(?![_\w])"),  # "check its id in Config", "the Filter of Config" (not "a list of PIs")
    re.compile(rf"^(?:{USER})(?: row\b|: (?!\d))"),                       # "Config: FlowLoadMode ...", "Workflows row 3: ..." (not "Projects: 0")
)
# Names and keys, not directions (their value or their sheet already says where to look); and the lines of the log, which support reads.
ALLOWED = {
    " | Statuses not in Workflows: ": "the name of a status field (modHealth, flow/health): its value names the real sheet",
    "Statuses not in Workflows": "the same field name, quoted by the manual",
    "Status not in Workflows": "the title of the data check DC08",
    "Status transition not in Transitions": "the title of the data check DC28",
    "(no PI in the sheet PIs)": "dashboard_spec.NO_PI: the label of a point of the PI axis and a key of PIMetrics, not a message",
    "all classified in Workflows": "reconciliation.CLASSIFIED: the assertion of a control, also a key of its formula",
    "no sheet Config": "a line of the log (modLog.LogParameters), read by support",
    "computed: sheet Workflows": "flow/raw_types: the source of a column, in the log only",
}
VBA_STRING = re.compile(r'"((?:[^"]|"")*)"')
SKIPPED = ("mockjira", "testbench", "mock_workbook.py")


def vba_literals(code: str) -> list[str]:
    out = []
    for line in code.splitlines():
        text, in_string = [], False
        for ch in line:                                                   # cut the comment: a ' outside a string
            if ch == '"':
                in_string = not in_string
            elif ch == "'" and not in_string:
                break
            text.append(ch)
        out += [m.replace('""', '"') for m in VBA_STRING.findall("".join(text))]
    return out


def python_literals(code: str) -> list[str]:
    tree = ast.parse(code)
    docstrings = {id(node.body[0].value) for node in ast.walk(tree)
                  if isinstance(node, (ast.Module, ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)) and node.body
                  and isinstance(node.body[0], ast.Expr) and isinstance(node.body[0].value, ast.Constant)}
    return [node.value for node in ast.walk(tree)
            if isinstance(node, ast.Constant) and isinstance(node.value, str) and id(node) not in docstrings]


def offenders(texts: list[str]) -> list[str]:
    return [t for t in texts if t not in ALLOWED and not any(a in t for a in ALLOWED) and any(p.search(t) for p in PATTERNS)]


def test_the_rule_catches_a_logical_name_and_lets_a_real_name_pass():
    assert offenders(["Sheet Iterations: row 3", "In the Config sheet check", "check its id in Config", "Config: FlowLoadMode must",
                      "Review: the sheets JiraFields and JiraStatuses", "Workflows row 4: no status"]) == [
        "Sheet Iterations: row 3", "In the Config sheet check", "check its id in Config", "Config: FlowLoadMode must",
        "Review: the sheets JiraFields and JiraStatuses", "Workflows row 4: no status"]
    assert offenders(["Sheet Iterations_from_user: row 3", "In the sheet ", " check Email", "Projects: 0 | Workflows: 0",
                      " | Workflows not read: 1", "Config_from_user row 2"]) == []
    assert vba_literals('x = "In the Config sheet" \' "the sheet PIs" in a comment') == ["In the Config sheet"]
    assert python_literals('"""the sheet PIs"""\nx = f"Sheet {a}: the Config sheet"') == ["Sheet ", ": the Config sheet"]


def test_every_message_of_the_vba_names_the_real_sheet():
    found = {path.name: offenders(vba_literals(path.read_text(encoding="ascii")))
             for path in sorted((ROOT / "vba").glob("*.bas")) if path.name != "modTests.bas"}
    assert {name: texts for name, texts in found.items() if texts} == {}


def test_every_message_of_the_python_package_names_the_real_sheet():
    found = {str(path.relative_to(ROOT)): offenders(python_literals(path.read_text(encoding="utf-8")))
             for path in sorted((ROOT / "src" / "jira2dashboard").rglob("*.py"))
             if not any(part in SKIPPED for part in path.parts)}
    assert {name: texts for name, texts in found.items() if texts} == {}
