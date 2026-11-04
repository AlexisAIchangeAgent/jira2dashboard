"""Guard: mappings (statuses, extract columns) are parameters written in the workbook, never VBA code.

Out of scope on purpose: Jira API identifiers (`"status"` in modFlow, JSON keys, modJira) and the fixtures of
modTests (they use French status names to test the code).
"""
import re
import unicodedata

import pytest

from jira2dashboard.params import model
from jira2dashboard.paths import VBA_DIR


def _ascii(text: str) -> str:
    """VBA source is ASCII: compare without accents."""
    return unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")


# What must not be written in code is what the parameters (params/*.md) define: the status names of the
# workflows and the columns of the ticket extract.
_SHEETS = {s.name: s for s in model.load_all()}
STATUS_NAMES = sorted({_ascii(row[2]) for row in _SHEETS["Workflows"].rows})
EXTRACT_COLUMNS = [row[0] for row in _SHEETS["FieldMapping"].rows]
# Same word, other meaning: `resp("Status")` is the HTTP status code key of the response dictionary
# (modHttp, modJira), not the Jira "Status" column.
# "Type" is the header of the Issues sheet written by modExtract (a sheet header, not a mapping).
HTTP_IDENTIFIERS = {"Status", "Type"}
_STRING = re.compile(r'"((?:[^"]|"")*)"')


def string_literals(vba: str) -> set[str]:
    """Every string literal of the code, comments removed."""
    found = set()
    for line in vba.splitlines():
        code = _strip_comment(line)
        found.update(m.group(1).replace('""', '"') for m in _STRING.finditer(code))
    return found


def _strip_comment(line: str) -> str:
    in_string = False
    for i, ch in enumerate(line):
        if ch == '"':
            in_string = not in_string
        elif ch == "'" and not in_string:
            return line[:i]
    return line


def _code_modules():
    return sorted(p for p in VBA_DIR.glob("*.bas") if p.stem != "modTests")


def test_string_literals_ignores_comments_and_unescapes_quotes():
    vba = 'x = "a""b" \' "commented"\n y = "c"'
    assert string_literals(vba) == {'a"b', "c"}


@pytest.mark.parametrize("module", _code_modules(), ids=lambda p: p.stem)
def test_no_status_name_in_vba_code(module):
    literals = string_literals(module.read_text(encoding="ascii"))
    assert not literals & set(STATUS_NAMES), sorted(literals & set(STATUS_NAMES))


@pytest.mark.parametrize("module", _code_modules(), ids=lambda p: p.stem)
def test_no_extract_column_name_in_vba_code(module):
    literals = string_literals(module.read_text(encoding="ascii"))
    hits = literals & (set(EXTRACT_COLUMNS) - HTTP_IDENTIFIERS)
    assert not hits, sorted(hits)
