"""The local log of the runs (vba/modLog.bas) must never receive the API token, an e-mail address, a URL, the JQL or an
answer of Jira: the calls to the logger pass step names, counts and saved error numbers only."""
import re

from jira2dashboard.params.model import load_all
from jira2dashboard.testbench import run_vba_tests as rvt

LOG_CALL = re.compile(r"\bLog(?:Line|Begin|End|Error|RunStart|RunEnd)\b(.*)$", re.MULTILINE)
SENSITIVE = re.compile(r"\bresp\b|\bcfg\b|\bmCfg\b|Token|Authorization|Email|BaseUrl|\burl\b|\bjql\b|\bmJql\b|\bbody\b|"
                       r"\.Text\b|\(\"Text\"\)", re.IGNORECASE)


def sensitive_log_calls(sources: dict[str, str]) -> list[tuple[str, str]]:
    return [(name, call.group(0).strip()) for name, code in sources.items() if name not in ("modLog", "modTests")
            for call in LOG_CALL.finditer(code) if SENSITIVE.search(call.group(1))]


def test_the_guard_catches_a_secret_passed_to_the_log():
    code = ('LogLine "HTTP", resp("Text")\nLogError "x", 1, cfg("Authorization")\nLogBegin "page " & mJql\n'
            'LogRunEnd "stopped " & url\n')
    assert len(sensitive_log_calls({"modX": code})) == 4
    assert sensitive_log_calls({"modX": 'LogBegin "search page " & (mPage + 1)\n'}) == []


def test_no_module_passes_a_secret_or_a_jira_answer_to_the_log():
    assert sensitive_log_calls(rvt.module_sources(rvt.ORDER)) == []


SHOWN = re.compile(r'Private Const SHOWN_PARAMETERS As String = ((?:"[^"]*"\s*&\s*_\s*\n\s*)*"[^"]*")')


def shown_parameters(source: str) -> set[str]:
    literal = "".join(re.findall(r'"([^"]*)"', SHOWN.search(source).group(1)))
    return {name for name in literal.split(",") if name}


def repo_settings() -> set[str]:
    """The parameters of Config with a value in params/config.md: the settings of the tool, the same for every client."""
    config = next(sheet for sheet in load_all() if sheet.name == "Config")
    value = config.header.index("Value")
    return {row[0] for row in config.rows if row[value] not in (None, "")}


def test_the_log_shows_in_clear_only_the_settings_that_have_a_value_in_the_repository():
    """modLog.LogParameters writes the inputs of a run (PARAM lines): the value only for the settings of the tool; every
    value typed at the client (connection, proxy, filter, field ids, AsOf) reads (set) or (empty) (modLog.ParameterShown)."""
    shown = shown_parameters(rvt.module_sources(rvt.ORDER)["modLog"])
    assert shown == repo_settings()
    assert not shown & {"BaseUrl", "Email", "ApiToken", "Filter", "ProxyServer", "ProxyBypass", "AsOf", "WorkTypeField"}
