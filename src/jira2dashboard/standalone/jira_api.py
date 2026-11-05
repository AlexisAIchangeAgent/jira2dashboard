"""The rules of the calls to Jira, pure (no network): what the VBA modules modJira and modHttp do, in Python.

The search: GET /rest/api/3/search/jql, 100 issues a page, paged by nextPageToken until isLast. The changelogs:
POST /rest/api/3/changelog/bulkfetch, 1000 keys a call, paged by nextPageToken. Retries after 429 and 5xx with a
growing wait, never shorter than Retry-After. The messages are those of the VBA, with the same error numbers.
"""
import base64
import json
from urllib.parse import quote

PAGE_SIZE = 100
CHANGELOG_CHUNK = 1000
MAX_SEARCH_PAGES = 2000
MAX_CHANGELOG_PAGES = 5000
RETRY_BASE_S, RETRY_MAX_S, RETRY_AFTER_MAX_S = 2.0, 30.0, 120.0
MAX_NETWORK_FAILURES = 3                  # status 0 (no answer): the third one stops the retries
WHITE = " \t\r\n "
ALWAYS_FIELDS = ("key", "project", "issuetype", "created", "status", "parent", "resolutiondate", "fixVersions",
                 "issuelinks")
FIELD_ORDER = ("StoryPointsField", "BusinessValueField", "PIField", "TeamField", "SprintField", "TargetStartField",
               "TargetEndField")                        # the order of modExtract.ExtractFields


class JiraError(Exception):
    """An error of the extraction, with the number of the VBA (vbObjectError + code)."""

    def __init__(self, code: int, message: str):
        super().__init__(message)
        self.code = code


def trim_white(value: str | None) -> str:
    return (value or "").strip(WHITE)


def _is_local(url: str) -> bool:
    """The fake Jira listens on this computer (or behind the fake proxy under a .invalid name): http is fine there."""
    lower = url.lower()
    if not lower.startswith("http://"):
        return False
    host = lower[7:].split("/", 1)[0].split("?", 1)[0].split(":", 1)[0]
    return host in ("127.0.0.1", "localhost") or host.endswith(".invalid")


def check_base_url(value: str) -> str:
    """The site without trailing slash. https only: the token would travel in clear text (http for the fake Jira)."""
    url = trim_white(value).rstrip("/")
    if url and not url.lower().startswith("https://") and not _is_local(url):
        raise JiraError(602, "BaseUrl must start with https:// (the API token would travel in clear text). "
                             "Fix it in the Config sheet.")
    return url


def normalize_proxy(value: str) -> str:
    """host:port whatever was typed: http://proxy.name.com:8080/ -> proxy.name.com:8080 (WinHTTP wants no scheme)."""
    proxy = trim_white(value)
    for scheme in ("http://", "https://"):
        if proxy.lower().startswith(scheme):
            proxy = proxy[len(scheme):]
    return proxy.rstrip("/")


def basic_auth(email: str, token: str) -> str:
    pair = f"{trim_white(email)}:{trim_white(token)}".encode("utf-8")
    return "Basic " + base64.b64encode(pair).decode("ascii")


def projects_jql(projects: list[tuple[str, str, str]], extra: str = "") -> str:
    """project in ("PRG","TA") [AND (filter)] ORDER BY key ASC, the Program project first (modJira.ProjectsJql)."""
    keys = [row[0] for row in sorted(projects, key=lambda r: (r[1] or "").lower() != "program") if row[0]]
    if not keys:
        raise JiraError(600, "No project to extract: fill the Projects sheet (Project, Role, Team).")
    jql = "project in (" + ",".join(f'"{key.strip(" ").replace(chr(34), "")}"' for key in keys) + ")"
    extra = trim_white(extra)
    if extra:
        if "order by" in extra.lower():
            raise JiraError(603, "The Filter of the Config sheet must not contain ORDER BY")
        jql += f" AND ({extra})"
    return jql + " ORDER BY key ASC"


def search_fields(work_type_field: str, fields: dict[str, str]) -> list[str]:
    """The fields the extraction asks for (modExtract.ExtractFields): the system ones, then the configured ids."""
    ids = [work_type_field] + [fields.get(name, "") for name in FIELD_ORDER]
    return list(ALWAYS_FIELDS) + [i for i in ids if i]


def encode(text: str) -> str:
    """UTF-8, every byte but 0-9 A-Z a-z - . _ ~ as %XX (modText.UrlEncode)."""
    return quote(text.encode("utf-8"), safe="")


def search_path(jql: str, fields: list[str], token: str = "") -> str:
    path = f"/rest/api/3/search/jql?jql={encode(jql)}&maxResults={PAGE_SIZE}&fields={encode(','.join(fields))}"
    return path + (f"&nextPageToken={encode(token)}" if token else "")


def changelog_body(keys: list[str], token: str = "") -> bytes:
    body: dict = {"issueIdsOrKeys": keys, "maxResults": CHANGELOG_CHUNK}
    if token:
        body["nextPageToken"] = token
    return json.dumps(body).encode("utf-8")


def chunks(keys: list[str], size: int = CHANGELOG_CHUNK) -> list[list[str]]:
    return [keys[i:i + size] for i in range(0, len(keys), size)]


def is_retryable(status: int) -> bool:
    return status == 429 or status >= 500


def retry_delay(attempt: int, retry_after: str | None, rnd: float) -> float:
    """Seconds before the next attempt: 2 s x 2^attempt (at most 30 s) with half jitter, or Retry-After (in seconds, at
    most 120 s) when it is longer. `rnd` is a random number in [0, 1)."""
    backoff = min(RETRY_BASE_S * 2 ** attempt, RETRY_MAX_S)
    backoff = backoff / 2 + backoff / 2 * rnd
    header = 0.0
    if retry_after:
        digits = retry_after.strip().split(".")[0]
        header = min(float(digits) + rnd * 0.5, RETRY_AFTER_MAX_S) if digits.isdecimal() else 0.0
    return max(header, backoff)


def jira_message(text: str) -> str:
    """What Jira says in an error answer (errorMessages and errors), else the start of the answer."""
    try:
        body = json.loads(text)
    except ValueError:
        return text.strip()[:500]
    if not isinstance(body, dict):
        return text.strip()[:500]
    messages = [str(m) for m in body.get("errorMessages") or []]
    messages += [f"{field}: {message}" for field, message in (body.get("errors") or {}).items()]
    return "; ".join(messages) or text.strip()[:500]


HINTS = {
    407: " Hint: the proxy asks for authentication. In the Config sheet check ProxyServer and ProxyWindowsAuth "
         "(Yes = the Windows session is sent to the proxy).",
    401: " Hint: Jira refused the account. In the Config sheet check Email and ApiToken.",
    0: " Hint: no answer. In the Config sheet check BaseUrl, ProxyServer and ProxyBypass (ProxyServer empty = the proxy "
       "of the Windows settings, DIRECT = no proxy).",
}


def status_error(status: int, attempts: int, error: str, text: str) -> JiraError:
    """The error of an answer that is not 200 (modJira.CheckStatus): status, attempts, what Jira said, a hint."""
    said = jira_message(text) if text else ""
    detail = " ".join(part for part in (error, said) if part)
    return JiraError(602, f"HTTP {status} after {attempts} attempt(s). {detail}{HINTS.get(status, '')}".replace("  ", " "))


def parse_json(text: str, what: str):
    """The answer as JSON; a proxy or a login page can answer 200 with HTML (modJira.RequireJson, error 607)."""
    if text.lstrip()[:1] not in ("{", "["):
        raise JiraError(607, f"Jira {what}: Jira answered 200 with something that is not JSON (a proxy or a login "
                             "page?). Check ProxyServer and BaseUrl in the Config sheet. Start of the answer: "
                             + text[:200])
    return json.loads(text)
