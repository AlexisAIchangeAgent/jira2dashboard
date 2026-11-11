"""The local log of the runs (vba/modLog.bas) must never receive the API token, an e-mail address, a URL, the JQL or an
answer of Jira: the calls to the logger pass step names, counts and saved error numbers only."""
import re

from jira2dashboard.params.model import CONFIG_PARAMETERS
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


MASKED = {"BaseUrl", "Email", "ApiToken", "ProxyServer", "ProxyBypass", "Filter"}     # the connection, the proxy, the JQL


def test_the_log_shows_every_parameter_in_clear_but_the_connection_the_proxy_and_the_filter():
    """modLog.LogParameters writes the inputs of a run (PARAM lines): the value of each parameter of Config (AsOf and the
    field ids too: AL's choice, 2026-11-11), but BaseUrl, Email, ApiToken, ProxyServer, ProxyBypass and Filter read (set) or
    (empty) (modLog.ParameterShown). A new parameter is masked until it is added to SHOWN_PARAMETERS."""
    shown = shown_parameters(rvt.module_sources(rvt.ORDER)["modLog"])
    assert shown == set(CONFIG_PARAMETERS) - MASKED
    assert not shown & MASKED
