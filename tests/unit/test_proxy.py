"""The proxy of the Python path: the order of modProxy.DetectProxy, then the real WinHTTP calls (ctypes structures)."""
from jira2dashboard.standalone import proxy as P
from jira2dashboard.testbench import run_vba_tests as rvt

URL = "https://site.atlassian.net"


def script(ok, found="", bypass="", error=0):
    calls = []

    def run(url, auto_detect, auto_config_url):
        calls.append((url, auto_detect, auto_config_url))
        return P.Script(ok, found, bypass, error)
    run.calls = calls
    return run


def test_the_proxy_script_comes_first():
    run = script(True, "proxy.pac.local:8080", "*.local")
    found = P.detect_proxy(URL, P.Settings(True, "http://wpad/proxy.pac", "fixed:3128", ""), run)
    assert found == P.Detected("proxy.pac.local:8080", "*.local", "proxy script (PAC): a proxy")
    assert run.calls == [(URL, True, "http://wpad/proxy.pac")]                       # one call, both ways


def test_wpad_can_say_direct():
    assert P.detect_proxy(URL, P.Settings(True, "", "", ""), script(True)) == P.Detected(
        "", "", "automatic detection (WPAD): direct")


def test_when_the_script_finds_nothing_the_fixed_proxy_then_direct():
    assert P.detect_proxy(URL, P.Settings(False, "http://x/p.pac", "fixed:3128", "<local>"), script(False, error=12167)) \
        == P.Detected("fixed:3128", "<local>",
                      "proxy script (PAC) found nothing (Windows error 12167), then the fixed proxy of the Windows settings")
    assert P.detect_proxy(URL, P.Settings(True, "", "", ""), script(False, error=12180)).mode == (
        "automatic detection (WPAD) found nothing (Windows error 12180), then direct (no proxy in the Windows settings)")


def test_no_setting_no_script():
    run = script(True, "never")
    assert P.detect_proxy(URL, P.Settings(False, "", "", ""), run) == P.Detected(
        "", "", "direct (no proxy in the Windows settings)")
    assert P.detect_proxy(URL, P.Settings(False, "", "fixed:1", "b"), run).server == "fixed:1" and run.calls == []


def test_config_decides_first():
    def detect(url):
        return P.Detected("http://detected:8080/", "*.detected", "proxy script (PAC): a proxy")
    assert P.choose_proxy(" direct ", "x", URL, detect) == P.Detected("", "x", "DIRECT in Config: no proxy")
    assert P.choose_proxy("http://typed:3128/", "", URL, detect) == P.Detected("typed:3128", "", "typed in Config")
    assert P.choose_proxy("", "", URL, detect) == P.Detected("detected:8080", "*.detected", "proxy script (PAC): a proxy")
    assert P.choose_proxy("", "*.mine", URL, detect).bypass == "*.mine"            # a typed bypass list stays


def test_the_real_settings_of_this_user_are_read():
    settings = P.windows_settings()
    assert isinstance(settings.auto_detect, bool) and all(isinstance(v, str) for v in settings[1:])
    assert P.detect_proxy(URL).mode                                                 # whatever this machine has


def test_winhttp_runs_a_real_proxy_script():
    with rvt.pac_server() as pac:
        found = P.script_proxy(URL, False, pac)
    assert found == P.Script(True, rvt.PAC_PROXY, "", 0)


def test_an_unreachable_script_is_a_windows_error():
    found = P.script_proxy(URL, False, "http://127.0.0.1:9/none.pac")
    assert not found.ok and found.error > 12000                                     # ERROR_WINHTTP_*
