"""The delivered VBA creates only two COM objects: the HTTP request and the dictionary.

ADODB.Stream (a binary stream next to an HTTP client) and MSXML2 (used as a Base64 codec) are what an endpoint protection reads
as a downloader. The text helpers are written in plain VBA; modTests keeps the two objects only as the oracle of its tests.
"""
import re

from jira2dashboard.build_workbook import MODULES
from jira2dashboard.testbench import run_vba_tests as rvt

ALLOWED = {"Scripting.Dictionary", "WinHttp.WinHttpRequest.5.1"}


def com_objects(sources: dict[str, str]) -> dict[str, set[str]]:
    """{module: the ProgIDs it creates with CreateObject}, comments left out."""
    found = {}
    for name, text in sources.items():
        code = "\n".join(line.split("'", 1)[0] for line in text.splitlines())
        found[name] = set(re.findall(r'CreateObject\("([^"]+)"\)', code))
    return found


def test_the_delivered_modules_create_only_the_http_request_and_the_dictionary():
    sources = rvt.module_sources(MODULES)
    assert len(sources) == len(MODULES) == 14
    extra = {name: ids - ALLOWED for name, ids in com_objects(sources).items() if ids - ALLOWED}
    assert not extra, f"COM objects an endpoint protection may block: {extra}"


def test_the_check_sees_an_object_in_code_and_ignores_a_comment():
    sources = {"modX": 'Set a = CreateObject("ADODB.Stream")\n\' Set b = CreateObject("MSXML2.DOMDocument.6.0")\n',
               "modY": 'Set d = CreateObject("Scripting.Dictionary")'}
    assert com_objects(sources) == {"modX": {"ADODB.Stream"}, "modY": {"Scripting.Dictionary"}}
