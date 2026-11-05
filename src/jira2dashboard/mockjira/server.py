"""HTTP server of the fake Jira (stdlib only). Runs in a thread on 127.0.0.1; no TLS (WinHTTP accepts http://).

    with MockJira(dataset) as jira:
        jira.base_url            # http://127.0.0.1:<port>  (the BaseUrl parameter of the workbook)
        jira.fail_next("POST", "/rest/api/3/changelog/bulkfetch", 429, times=2, retry_after="1")
        jira.requests            # what the client sent

The same fault injection is reachable over HTTP, so that VBA tests can trigger it:
    POST /__mock__/fail   {"method":"GET","path":"/rest/api/3/search/jql","status":429,"times":2,"retryAfter":"0"}
    POST /__mock__/reset  clears the faults and the request log
    GET /proxy.pac        a proxy script (PAC) that sends every address to this fake as a proxy (no account asked)
"""
import base64
import json
import re
import select
import socket
import threading
import time
import urllib.parse
from dataclasses import dataclass, field
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from jira2dashboard.mockjira import api, config
from jira2dashboard.mockjira.dataset import Dataset
from jira2dashboard.mockjira.errors import ApiError

GONE_PATHS = ("/rest/api/3/search", "/rest/api/3/search/id")
PROJECT_STATUSES = re.compile(r"/rest/api/3/project/([^/]+)/statuses")
GONE_MESSAGE = "The requested API has been removed. Please migrate to the /rest/api/3/search/jql API."
PAC_PATH = "/proxy.pac"                       # a proxy script (PAC) that sends every address to this fake as a proxy
PAC_TYPE = "application/x-ns-proxy-autoconfig"
MOCK_EMAIL = "mock@example.com"                # the account of `run.ps1 mock-jira` and of the manual-test copy
MOCK_TOKEN = "mock-token"


def pac_script(proxy: str) -> str:
    """A proxy script (PAC) that sends every address to `proxy` (host:port): the Windows setting in the tests of the
    proxy detection (VBA M_AutoProxyPac with this fake, T_ProxyScript with run_vba_tests.pac_server)."""
    return f'function FindProxyForURL(url, host) {{ return "PROXY {proxy}"; }}\n'


@dataclass
class Fault:
    method: str
    path: str
    status: int
    remaining: int
    retry_after: str | None = None
    body: object = None


@dataclass(frozen=True)
class Request:
    method: str
    path: str
    query: dict
    body: str
    headers: dict


@dataclass
class _State:
    dataset: Dataset
    credentials: tuple[str, str] | None
    tz_minutes: int
    page_clamp: int
    bulk_clamp: int
    near_limit: bool
    proxy_challenge: str | None = None
    delay: float = 0.0                       # seconds every API call waits before it answers (slow Jira)
    base_url: str = ""
    address: tuple = ("", 0)
    faults: list[Fault] = field(default_factory=list)
    requests: list[Request] = field(default_factory=list)
    lock: threading.Lock = field(default_factory=threading.Lock)


MAX_BODY = 10_000_000                       # bytes a request may carry


def _json_params(body: dict) -> dict:
    """POST /search/jql takes fields as a list and expand as a string; the GET form takes comma-separated text."""
    return {"jql": body.get("jql"), "fields": body.get("fields") or [], "expand": body.get("expand"),
            "maxResults": body.get("maxResults"), "nextPageToken": body.get("nextPageToken")}


def _query_params(query: dict) -> dict:
    first = lambda name: query.get(name, [None])[0]            # noqa: E731
    fields = [f for f in (first("fields") or "").split(",") if f]
    return {"jql": first("jql"), "fields": fields, "expand": first("expand"),
            "maxResults": first("maxResults"), "nextPageToken": first("nextPageToken")}


def _make_handler(state: _State):
    class Handler(BaseHTTPRequestHandler):
        protocol_version = "HTTP/1.0"

        def log_message(self, *args):                          # keep the test output quiet
            pass

        def do_GET(self):
            self._serve()

        def do_POST(self):
            self._serve()

        def do_CONNECT(self):
            """HTTPS through a proxy starts with CONNECT host:port. The tunnel leads back to this server
            (plain HTTP inside: no TLS here)."""
            with state.lock:
                state.requests.append(Request("CONNECT", self.path, {}, "",
                                              {k.lower(): v for k, v in self.headers.items()}))
            if self._proxy_auth_missing():
                self._send(407, {"errorMessages": ["Proxy Authentication Required"], "errors": {}},
                           {"Proxy-Authenticate": state.proxy_challenge})
                return
            self.send_response(200, "Connection Established")
            self.end_headers()
            self.close_connection = True
            with socket.create_connection(state.address) as upstream:
                self._relay(self.connection, upstream)

        @staticmethod
        def _relay(client, upstream):
            sockets = [client, upstream]
            while True:
                readable, _, _ = select.select(sockets, [], [], 10)
                if not readable:
                    return
                for source in readable:
                    try:
                        data = source.recv(65536)
                        if not data:
                            return
                        (upstream if source is client else client).sendall(data)
                    except OSError:                  # a connection reset by the peer ends the tunnel quietly
                        return

        def _proxy_auth_missing(self) -> bool:
            """With a proxy challenge set, a request that came through the proxy needs Proxy-Authorization.
            Any value is accepted: a real NTLM / Kerberos token cannot be checked here."""
            return bool(state.proxy_challenge) and not self.headers.get("Proxy-Authorization")

        def _send(self, status: int, body, headers: dict | None = None):
            headers = dict(headers or {})
            if isinstance(body, str):
                raw, kind = body.encode("utf-8"), "text/plain;charset=UTF-8"
            else:
                raw, kind = json.dumps(body, ensure_ascii=False).encode("utf-8"), "application/json;charset=UTF-8"
            self.send_response(status)
            self.send_header("Content-Type", headers.pop("Content-Type", kind))
            self.send_header("Content-Length", str(len(raw)))
            if state.near_limit:
                self.send_header("X-RateLimit-NearLimit", "true")
            for name, value in (headers or {}).items():
                self.send_header(name, value)
            self.end_headers()
            self.wfile.write(raw)

        def _serve(self):
            parsed = urllib.parse.urlsplit(self.path)
            query = urllib.parse.parse_qs(parsed.query)
            raw_length = self.headers.get("Content-Length") or "0"
            if not raw_length.isdecimal() or int(raw_length) > MAX_BODY:          # not a number, negative or huge
                self._send(400, {"errorMessages": ["bad Content-Length"], "errors": {}}, {})
                return
            length = int(raw_length)
            text = self.rfile.read(length).decode("utf-8") if length else ""
            with state.lock:
                state.requests.append(Request(self.command, parsed.path, query, text,
                                              {k.lower(): v for k, v in self.headers.items()}))
            if state.delay:
                time.sleep(state.delay)
            try:
                status, body, headers = self._dispatch(parsed.path, query, text)
            except ApiError as err:
                status, body, headers = err.status, err.body, err.headers
            except Exception as err:                            # noqa: BLE001 - a mock must tell, not hide
                status, body, headers = 500, {"errorMessages": [f"mock error: {type(err).__name__}: {err}"], "errors": {}}, {}
            self._send(status, body, headers)

        def _dispatch(self, path: str, query: dict, text: str):
            if path.startswith("/__mock__/"):
                return self._control(path, text)
            if (self.command, path) == ("GET", PAC_PATH) and not self.path.startswith(("http://", "https://")):
                return 200, pac_script(state.base_url.removeprefix("http://")), {"Content-Type": PAC_TYPE}    # Windows asks with no account
            if self.path.startswith(("http://", "https://")) and self._proxy_auth_missing():    # came through the proxy
                return 407, {"errorMessages": ["Proxy Authentication Required"], "errors": {}}, \
                    {"Proxy-Authenticate": state.proxy_challenge}
            if state.credentials and self.headers.get("Authorization") != _basic(*state.credentials):
                return 401, "Unauthorized", {}
            fault = self._take_fault(path)
            if fault:
                headers = {"Retry-After": fault.retry_after} if fault.retry_after is not None else {}
                if fault.status == 429:
                    headers["RateLimit-Reason"] = "jira-burst-based"
                body = fault.body if fault.body is not None else {"errorMessages": [f"mock fault {fault.status}"], "errors": {}}
                return fault.status, body, headers
            if path in GONE_PATHS:
                return 410, {"errorMessages": [GONE_MESSAGE], "errors": {}}, {}
            return self._route(path, query, text)

        def _route(self, path: str, query: dict, text: str):
            ds, tz = state.dataset, state.tz_minutes
            method = self.command
            body = json.loads(text) if text else {}
            if (method, path) == ("GET", "/rest/api/3/search/jql"):
                return 200, api.search(ds, _query_params(query), tz=tz, base_url=state.base_url,
                                       page_clamp=state.page_clamp), {}
            if (method, path) == ("POST", "/rest/api/3/search/jql"):
                return 200, api.search(ds, _json_params(body), tz=tz, base_url=state.base_url,
                                       page_clamp=state.page_clamp), {}
            if (method, path) == ("POST", "/rest/api/3/search/approximate-count"):
                return 200, api.approximate_count(ds, body), {}
            if (method, path) == ("POST", "/rest/api/3/changelog/bulkfetch"):
                return 200, api.bulk_changelog(ds, body, tz=tz, bulk_clamp=state.bulk_clamp), {}
            if (method, path) == ("GET", "/rest/api/3/field"):
                return 200, api.field_catalog(ds), {}
            if (method, path) == ("GET", "/rest/api/3/myself"):
                return 200, api.myself(), {}
            if (method, path) == ("GET", "/rest/api/3/project/search"):
                return 200, config.project_search(ds, query), {}
            statuses = PROJECT_STATUSES.fullmatch(path)
            if method == "GET" and statuses:
                return 200, config.project_statuses(ds, urllib.parse.unquote(statuses.group(1))), {}
            if (method, path) == ("POST", "/rest/api/3/workflows/preview"):
                return 200, config.workflows_preview(ds, body), {}
            known = {"/rest/api/3/search/jql", "/rest/api/3/search/approximate-count",
                     "/rest/api/3/changelog/bulkfetch", "/rest/api/3/field", "/rest/api/3/myself",
                     "/rest/api/3/project/search", "/rest/api/3/workflows/preview"}
            if path in known:
                return 405, {"errorMessages": [f"Method {method} is not allowed on {path}."], "errors": {}}, {}
            return 404, {"errorMessages": ["Not Found"], "errors": {}}, {}

        def _take_fault(self, path: str) -> Fault | None:
            with state.lock:
                for fault in state.faults:
                    if fault.method == self.command and fault.path == path and fault.remaining > 0:
                        fault.remaining -= 1
                        return fault
            return None

        def _control(self, path: str, text: str):
            if path == "/__mock__/reset":
                with state.lock:
                    state.faults.clear()
                    state.requests.clear()
                return 200, {"ok": True}, {}
            if path == "/__mock__/fail":
                spec = json.loads(text)
                with state.lock:
                    state.faults.append(Fault(spec["method"], spec["path"], int(spec["status"]),
                                              int(spec.get("times", 1)), spec.get("retryAfter"), spec.get("body")))
                return 200, {"ok": True}, {}
            return 404, {"errorMessages": ["Not Found"], "errors": {}}, {}

    return Handler


def _basic(email: str, token: str) -> str:
    return "Basic " + base64.b64encode(f"{email}:{token}".encode("utf-8")).decode("ascii")


class _ExclusiveServer(ThreadingHTTPServer):
    allow_reuse_address = False              # on Windows SO_REUSEADDR lets a second server take the port of a running one


class MockJira:
    """The fake Jira. `page_clamp` / `bulk_clamp` cap what one response can hold (Jira serves smaller pages than
    asked when many fields are requested: the client must read what came back). `proxy_challenge` ("Negotiate",
    "NTLM", "Basic realm=x") makes it behave as a proxy that wants authentication: requests that came through
    the proxy (absolute form or CONNECT) get 407 until they carry a Proxy-Authorization header."""

    def __init__(self, dataset: Dataset, *, host: str = "127.0.0.1", port: int = 0,
                 credentials: tuple[str, str] | None = None, tz_minutes: int = 120,
                 page_clamp: int = 100, bulk_clamp: int = 1000, near_limit: bool = False,
                 proxy_challenge: str | None = None, delay: float = 0.0):
        self._state = _State(dataset, credentials, tz_minutes, page_clamp, bulk_clamp, near_limit, proxy_challenge,
                             delay=delay)
        self._server = _ExclusiveServer((host, port), _make_handler(self._state))
        self._state.address = self._server.server_address[:2]
        self._server.daemon_threads = True
        self._thread = threading.Thread(target=self._server.serve_forever, kwargs={"poll_interval": 0.02}, daemon=True)
        self._state.base_url = self.base_url

    @property
    def base_url(self) -> str:
        host, port = self._server.server_address[:2]
        return f"http://{host}:{port}"

    @property
    def dataset(self) -> Dataset:
        return self._state.dataset

    @property
    def requests(self) -> list[Request]:
        with self._state.lock:
            return list(self._state.requests)

    def fail_next(self, method: str, path: str, status: int, *, times: int = 1, retry_after: str | None = None,
                  body=None) -> None:
        """The next `times` calls of method + path answer with `status` (429 and 503 usually carry Retry-After)."""
        with self._state.lock:
            self._state.faults.append(Fault(method, path, status, times, retry_after, body))

    def start(self) -> "MockJira":
        self._thread.start()
        return self

    def stop(self) -> None:
        if self._thread.is_alive():          # shutdown() waits for a serve loop: it would block if start() was never called
            self._server.shutdown()
        self._server.server_close()
        self._thread.join(timeout=5)

    def __enter__(self) -> "MockJira":
        return self.start()

    def __exit__(self, *exc) -> None:
        self.stop()
