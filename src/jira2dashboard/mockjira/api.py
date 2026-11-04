"""The responses of the fake Jira, as plain dicts. Pure: no HTTP here (see server.py).

Shapes follow docs/jira-cloud-api-vba-reference.md section 1 and the Jira Cloud platform REST API v3.
"""
import base64
import datetime as dt
import hashlib
import json

from jira2dashboard.mockjira import jql
from jira2dashboard.mockjira.dataset import Dataset, Issue
from jira2dashboard.mockjira.errors import ApiError

DEFAULT_PAGE = 50                   # /search/jql default maxResults
BULK_ISSUES_MAX = 1000              # changelog/bulkfetch: max issues per call
BULK_FIELD_IDS_MAX = 10


def jira_time(value: dt.datetime, tz_minutes: int) -> str:
    """Jira style: 2026-10-06T10:00:00.000+0200 (the offset has no colon)."""
    local = value + dt.timedelta(minutes=tz_minutes)
    sign = "+" if tz_minutes >= 0 else "-"
    hours, minutes = divmod(abs(tz_minutes), 60)
    return f"{local:%Y-%m-%dT%H:%M:%S}.{local.microsecond // 1000:03d}{sign}{hours:02d}{minutes:02d}"


def _error(status: int, message: str, **kwargs) -> ApiError:
    return ApiError(status, {"errorMessages": [message], "errors": {}}, **kwargs)


# ---------------------------------------------------------------- issues

def _sprint_json(sprint, tz: int) -> dict:
    return {"id": sprint.id, "name": sprint.name, "state": sprint.state, "boardId": 1,
            "startDate": jira_time(sprint.start, tz), "endDate": jira_time(sprint.end, tz)}


def _links_json(issue: Issue, ds: Dataset) -> list[dict]:
    """The links of the issue, as Jira shows them on both ends: "blocks" outward here, "is blocked by" inward there."""
    kind = {"id": "10000", "name": "Blocks", "inward": "is blocked by", "outward": "blocks"}
    links = [{"id": f"{issue.id}-out-{n}", "type": kind, "outwardIssue": {"key": key}} for n, key in enumerate(issue.blocks)]
    for other in ds.issues:
        if issue.key in other.blocks:
            links.append({"id": f"{other.id}-in-{issue.id}", "type": kind, "inwardIssue": {"key": other.key}})
    return links


def _all_fields(issue: Issue, ds: Dataset, tz: int) -> dict:
    catalog, ids = ds.catalog, ds.fields
    key, category = catalog.jira_category(issue.workflow, issue.status)
    resolution = ds.resolution(issue)
    return {
        "summary": issue.summary,
        "status": {"id": catalog.ids.get(issue.status, "0"), "name": issue.status,
                   "statusCategory": {"key": key, "name": category}},
        "issuetype": {"name": issue.type},
        "project": {"key": issue.project},
        "created": jira_time(issue.created, tz),
        "updated": jira_time(issue.updated, tz),
        "resolution": {"name": resolution[0]} if resolution else None,
        "resolutiondate": jira_time(resolution[1], tz) if resolution else None,
        "duedate": issue.due.isoformat() if issue.due else None,
        "labels": list(issue.labels),
        "fixVersions": [{"name": v} for v in issue.fix_versions],
        "parent": {"key": issue.parent} if issue.parent else None,
        "priority": {"name": issue.priority},
        "assignee": {"displayName": issue.assignee, "accountId": f"acc-{issue.assignee.split()[0].lower()}"}
        if issue.assignee else None,
        ids.story_points: issue.story_points,
        ids.work_type: issue.work_type,
        ids.pi: issue.pi,
        ids.business_value: issue.business_value,
        ids.team: {"id": f"team-{issue.team.lower()}", "name": issue.team} if issue.team else None,
        ids.sprint: [_sprint_json(s, tz) for s in issue.sprints] or None,
        ids.target_start: issue.target_start.isoformat() if issue.target_start else None,
        ids.target_end: issue.target_end.isoformat() if issue.target_end else None,
        "issuelinks": _links_json(issue, ds),
    }


def issue_json(issue: Issue, ds: Dataset, fields: list[str], tz: int, base_url: str) -> dict:
    """An issue with the requested fields. `*all` = every field; no field requested = the id only (like Jira)."""
    if not fields:
        return {"id": issue.id}
    everything = _all_fields(issue, ds, tz)
    wanted = everything if "*all" in fields else {f: everything[f] for f in fields if f in everything}
    return {"id": issue.id, "self": f"{base_url}/rest/api/3/issue/{issue.id}", "key": issue.key, "fields": wanted}


def _names(ds: Dataset, fields: list[str]) -> dict:
    catalog = {f["id"]: f["name"] for f in ds.fields.catalog()}
    return {fid: name for fid, name in catalog.items() if "*all" in fields or fid in fields}


# ---------------------------------------------------------------- paging

def _encode_token(offset: int, fingerprint: str) -> str:
    return base64.urlsafe_b64encode(f"{offset}:{fingerprint}".encode()).decode()


def _decode_token(token: str, fingerprint: str) -> int:
    try:
        offset, seen = base64.urlsafe_b64decode(token.encode()).decode().split(":", 1)
        if seen != fingerprint or int(offset) < 0:
            raise ValueError
        return int(offset)
    except (ValueError, UnicodeDecodeError):
        raise _error(400, "The nextPageToken is invalid or does not belong to this query.") from None


def _fingerprint(*parts) -> str:
    return hashlib.sha1(json.dumps(parts, sort_keys=True).encode()).hexdigest()[:10]


def _page_size(requested, default: int, clamp: int) -> int:
    """The size actually served: what was asked (or the default), never above the server clamp."""
    try:
        size = int(requested) if requested not in (None, "") else default
    except (TypeError, ValueError):
        raise _error(400, f"maxResults must be a number, got {requested!r}") from None
    if size < 1:
        raise _error(400, "maxResults must be at least 1.")
    return min(size, clamp)


# ---------------------------------------------------------------- endpoints

def _select(ds: Dataset, query: jql.Query) -> list[Issue]:
    issues = [i for i in ds.issues if jql.matches(i, query, ds.as_of)]
    sort_key = {"key": lambda i: (i.project, i.number), "created": lambda i: (i.created, i.project, i.number),
                "updated": lambda i: (i.updated, i.project, i.number)}[query.order_by]
    return sorted(issues, key=sort_key, reverse=query.descending)


def search(ds: Dataset, params: dict, *, tz: int, base_url: str, page_clamp: int) -> dict:
    """GET or POST /rest/api/3/search/jql. params: jql, fields (list), expand, maxResults, nextPageToken."""
    query = jql.parse(params.get("jql", ""))
    fields = params.get("fields") or []
    expand = [e.strip() for e in (params.get("expand") or "").split(",") if e.strip()]
    issues = _select(ds, query)
    fingerprint = _fingerprint(params.get("jql"))
    token = params.get("nextPageToken")
    offset = _decode_token(token, fingerprint) if token else 0
    size = _page_size(params.get("maxResults"), DEFAULT_PAGE, page_clamp)
    page = issues[offset:offset + size]
    body: dict = {"issues": [issue_json(i, ds, fields, tz, base_url) for i in page]}
    if "names" in expand and fields:
        body["names"] = _names(ds, fields)
    more = offset + size < len(issues)
    body["nextPageToken"] = _encode_token(offset + size, fingerprint) if more else None   # null on the last page
    body["isLast"] = not more
    return body


def approximate_count(ds: Dataset, body: dict) -> dict:
    """POST /rest/api/3/search/approximate-count."""
    return {"count": len(_select(ds, jql.parse(body.get("jql", ""))))}


class _KeyMove:
    """A change of key (the issue moved to another project): author and id like a status change."""

    def __init__(self, id: str, at: dt.datetime):
        self.id, self.at, self.author, self.author_id = id, at, "Ann Martin", "acc-ann"


def _key_item(issue: Issue) -> dict:
    return {"field": "Key", "fieldtype": "jira", "fieldId": "issuekey", "from": None,
            "fromString": f"OLD-{issue.number}", "to": None, "toString": issue.key}


def bulk_changelog(ds: Dataset, body: dict, *, tz: int, bulk_clamp: int) -> dict:
    """POST /rest/api/3/changelog/bulkfetch: the status histories, oldest change first, paged by token."""
    refs = body.get("issueIdsOrKeys")
    if not isinstance(refs, list) or not refs:
        raise _error(400, "issueIdsOrKeys is required and must not be empty.")
    if len(refs) > BULK_ISSUES_MAX:
        raise _error(400, f"issueIdsOrKeys can hold at most {BULK_ISSUES_MAX} issues, got {len(refs)}.")
    field_ids = body.get("fieldIds") or []
    if len(field_ids) > BULK_FIELD_IDS_MAX:
        raise _error(400, f"fieldIds can hold at most {BULK_FIELD_IDS_MAX} ids, got {len(field_ids)}.")
    issues = [i for i in (ds.by_key_or_id(str(r)) for r in refs) if i]
    histories = [(i, c) for i in issues for c in i.changes] if not field_ids or "status" in field_ids else []
    if not field_ids or "issuekey" in field_ids:        # moves between projects: the key changed
        histories += [(i, _KeyMove(f"{i.id}k{n}", at)) for i in issues for n, at in enumerate(i.key_changes, start=1)]
    histories.sort(key=lambda pair: (pair[1].at, int(pair[0].id), pair[1].id))   # by date, then smallest issue id
    fingerprint = _fingerprint(sorted(str(r) for r in refs), field_ids)
    token = body.get("nextPageToken")
    offset = _decode_token(token, fingerprint) if token else 0
    size = _page_size(body.get("maxResults"), bulk_clamp, bulk_clamp)
    grouped: dict[str, list] = {}
    for issue, change in histories[offset:offset + size]:
        grouped.setdefault(issue.id, []).append({
            "id": change.id, "created": jira_time(change.at, tz),
            "author": {"accountId": change.author_id, "displayName": change.author},
            "items": [_key_item(issue) if isinstance(change, _KeyMove) else
                      {"field": "status", "fieldtype": "jira", "fieldId": "status",
                       "from": ds.catalog.ids.get(change.from_status, "0"), "fromString": change.from_status,
                       "to": ds.catalog.ids.get(change.to_status, "0"), "toString": change.to_status}],
        })
    out: dict = {"issueChangeLogs": [{"issueId": iid, "changeHistories": hs} for iid, hs in grouped.items()]}
    more = offset + size < len(histories)
    out["nextPageToken"] = _encode_token(offset + size, fingerprint) if more else None    # null on the last page
    return out


def field_catalog(ds: Dataset) -> list[dict]:
    """GET /rest/api/3/field."""
    return ds.fields.catalog()


def myself() -> dict:
    """GET /rest/api/3/myself (credentials check)."""
    return {"accountId": "acc-mock", "displayName": "Mock User", "emailAddress": "mock@example.com", "active": True}
