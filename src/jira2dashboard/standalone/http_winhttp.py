"""The calls to Jira with WinHTTP, the network stack of the VBA (modHttp): a proxy that asks for the Windows session
is answered with it (ProxyWindowsAuth), the timeouts and the retries are those of Config.

The answer is read from ResponseBody and decoded as UTF-8 whatever the headers say (ResponseText would turn
"Terminé" into mojibake). Every failed attempt is logged with what Jira said; the wait between two attempts is
injectable, so the tests never sleep.
"""
import logging
import random
import time
from typing import Callable
from urllib.parse import urlsplit

from jira2dashboard.standalone import jira_api as J

PROXY_SETTING_PROXY = 2
AUTOLOGON_ALWAYS, AUTOLOGON_NEVER = 0, 2
RESOLVE_MS = 10_000
log = logging.getLogger("jira2dashboard.extract")


class WinHttp:
    """One request through the COM object WinHttp.WinHttpRequest.5.1 (synchronous: no Excel to keep alive here)."""

    def __init__(self, timeout_s: int, proxy_server: str, proxy_bypass: str, windows_auth: bool):
        self.timeout_ms = timeout_s * 1000
        self.proxy_server, self.proxy_bypass, self.windows_auth = proxy_server, proxy_bypass, windows_auth

    def send(self, method: str, url: str, headers: dict[str, str], body: bytes | None):
        """(status, answer bytes, Retry-After, error). Status 0 = no answer (the error says why)."""
        import pywintypes
        import win32com.client
        try:
            request = win32com.client.Dispatch("WinHttp.WinHttpRequest.5.1")
            request.Open(method, url, False)
            if self.proxy_server:
                request.SetProxy(PROXY_SETTING_PROXY, self.proxy_server, self.proxy_bypass)
            if not self.windows_auth:
                request.SetAutoLogonPolicy(AUTOLOGON_NEVER)
            elif self.proxy_server:
                request.SetAutoLogonPolicy(AUTOLOGON_ALWAYS)
            request.SetTimeouts(RESOLVE_MS, self.timeout_ms, self.timeout_ms, self.timeout_ms)
            for name, value in headers.items():
                request.SetRequestHeader(name, value)
            if body is None:
                request.Send()
            else:
                request.Send(body)
            status = int(request.Status)
            raw = request.ResponseBody
            data = bytes(raw) if raw is not None else b""
            try:
                retry_after = str(request.GetResponseHeader("Retry-After"))
            except pywintypes.com_error:
                retry_after = ""
            return status, data, retry_after, ""
        except pywintypes.com_error as err:
            info = err.excepinfo or (None,) * 6                 # the WinHTTP error is inside DISP_E_EXCEPTION
            code = info[5] or err.hresult
            detail = info[2] or err.strerror
            return 0, b"", "", f"{code & 0xFFFFFFFF:#010x} {str(detail or '').strip()}"


class JiraClient:
    """Search and changelogs of Jira with the retries of the VBA (429 and 5xx; three network failures stop)."""

    def __init__(self, base_url: str, authorization: str, max_retries: int, transport,
                 sleep: Callable[[float], None] = time.sleep, rnd: Callable[[], float] = random.random):
        self.base_url, self.max_retries, self.transport = base_url, max_retries, transport
        self.headers = {"Accept": "application/json", "Accept-Charset": "utf-8", "Authorization": authorization}
        self.sleep, self.rnd = sleep, rnd
        self.retries = 0                                        # attempts that were not the last: for the Run sheet

    def call(self, method: str, path: str, what: str, body: bytes | None = None):
        headers = dict(self.headers)
        if body is not None:
            headers["Content-Type"] = "application/json; charset=utf-8"
        network_failures, attempts = 0, 0
        status, text, error = 0, "", ""
        for attempt in range(self.max_retries + 1):
            attempts += 1
            status, data, retry_after, error = self.transport.send(method, self.base_url + path, headers, body)
            text = data.decode("utf-8", errors="replace")
            if status == 200:
                break
            said = error if status == 0 else J.jira_message(text)
            last = attempt == self.max_retries
            if status == 0:
                network_failures += 1
                last = last or network_failures >= J.MAX_NETWORK_FAILURES
            elif not J.is_retryable(status):
                last = True
            wait = 0.0 if last else J.retry_delay(attempt, retry_after, self.rnd())
            log.warning("%s: %s %s -> HTTP %s (attempt %d of %d)%s%s. %s", what, method, urlsplit(path).path, status,
                        attempts, self.max_retries + 1, f", Retry-After {retry_after}" if retry_after else "",
                        "" if last else f", retry in {wait:.1f} s", said)
            if last:
                break
            self.retries += 1
            self.sleep(wait)
        if status != 200:
            raise J.status_error(status, attempts, error, text)
        log.debug("%s: %s %s -> HTTP 200 (%d bytes, attempt %d)", what, method, urlsplit(path).path, len(text), attempts)
        return J.parse_json(text, what)

    def search_all(self, jql: str, fields: list[str]) -> list[dict]:
        """Every issue of the JQL, page after page (nextPageToken until isLast)."""
        issues, token = [], ""
        for page in range(1, J.MAX_SEARCH_PAGES + 1):
            answer = self.call("GET", J.search_path(jql, fields, token), "search")
            issues += answer.get("issues") or []
            log.info("END search page %d: %d issues", page, len(issues))
            following = answer.get("nextPageToken") or ""
            if answer.get("isLast") is True or not following:
                return issues
            if following == token:
                raise J.JiraError(606, "Jira keeps sending the same page token: stop here and try again later.")
            token = following
        raise J.JiraError(601, f"Jira search did not end after {J.MAX_SEARCH_PAGES} pages: narrow the extraction "
                               "with the Filter of the Config sheet.")

    def changelogs(self, keys: list[str]) -> list[dict]:
        """The changelog pages of the issues, 1000 keys a call."""
        pages = []
        for number, chunk in enumerate(J.chunks(keys)):
            token, seen = "", 0
            while True:
                answer = self.call("POST", "/rest/api/3/changelog/bulkfetch", "changelogs",
                                   J.changelog_body(chunk, token))
                pages.append(answer)
                seen += 1
                following = answer.get("nextPageToken") or ""
                if not following:
                    break
                if following == token or seen > J.MAX_CHANGELOG_PAGES:
                    raise J.JiraError(505, "Jira keeps sending the same changelog page: stop here and try again later.")
                token = following
            log.info("END changelogs from issue %d of %d", number * J.CHANGELOG_CHUNK + 1, len(keys))
        return pages
