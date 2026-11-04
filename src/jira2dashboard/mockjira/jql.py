"""The part of JQL that the extraction needs. Pure.

    project in ("A","B") [AND (<restriction> [AND <restriction>]...)] [ORDER BY key|created|updated [ASC|DESC]]

A restriction is `created` / `updated` (>=, >, <=, < a date "2026-10-01" or a relative `-30d`, `-2w`),
`issuetype` or `status` (= X, or in (X, Y)). Like Jira, an unbounded query (no project) is refused. Anything
else is refused with a clear message instead of being silently misread.
"""
import datetime as dt
import re
from dataclasses import dataclass

from jira2dashboard.mockjira.errors import ApiError

ORDER_FIELDS = ("key", "created", "updated")
_ORDER = re.compile(r"\s*\border\s+by\s+(\w+)(?:\s+(asc|desc))?\s*$", re.IGNORECASE)
_AND = re.compile(r"\s+AND\s+", re.IGNORECASE)
_PROJECT_IN = re.compile(r"project\s+in\s*\(([^)]*)\)", re.IGNORECASE)
_PROJECT_EQ = re.compile(r"project\s*=\s*(\"[^\"]+\"|\w+)", re.IGNORECASE)
_NAME_IN = re.compile(r"(issuetype|status)\s+in\s*\(([^)]*)\)", re.IGNORECASE)
_NAME_EQ = re.compile(r"(issuetype|status)\s*=\s*(\"[^\"]+\"|[\w-]+)", re.IGNORECASE)
_DATE = re.compile(r"(created|updated)\s*(>=|<=|>|<)\s*(\"?)(\d{4}-\d{2}-\d{2}|-\d+[dw])\3", re.IGNORECASE)


@dataclass(frozen=True)
class Condition:
    """field: issuetype / status (`values` = accepted names) or created / updated (`op`, `bound`)."""
    field: str
    values: tuple[str, ...] = ()
    op: str = ""
    bound: str = ""                  # "2026-10-01" or "-30d"


@dataclass(frozen=True)
class Query:
    projects: tuple[str, ...]
    order_by: str = "key"
    descending: bool = False
    conditions: tuple[Condition, ...] = ()


def _bad(message: str) -> ApiError:
    return ApiError(400, {"errorMessages": [message], "errors": {}})


def _unquote(items: str) -> tuple[str, ...]:
    return tuple(p.strip().strip('"') for p in items.split(",") if p.strip())


def _split_and(text: str) -> list[str]:
    """Split on AND outside parentheses and quotes."""
    parts, depth, quoted, start, i = [], 0, False, 0, 0
    while i < len(text):
        ch = text[i]
        if ch == '"':
            quoted = not quoted
        elif not quoted:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif depth == 0 and (m := _AND.match(text, i)):
                parts.append(text[start:i])
                i, start = m.end(), m.end()
                continue
        i += 1
    parts.append(text[start:])
    return [p.strip() for p in parts if p.strip()]


def _wrapped(text: str) -> bool:
    """True if the first '(' closes at the very end: (a AND b)."""
    if not (text.startswith("(") and text.endswith(")")):
        return False
    depth = 0
    for i, ch in enumerate(text):
        depth += ch == "("
        depth -= ch == ")"
        if depth == 0 and i < len(text) - 1:
            return False
    return True


def _flatten(text: str) -> list[str]:
    clauses = []
    for part in _split_and(text):
        while _wrapped(part):
            part = part[1:-1].strip()
        clauses.extend(_flatten(part) if len(_split_and(part)) > 1 else [part])
    return clauses


def parse(jql: str) -> Query:
    text = (jql or "").strip()
    order_by, descending = "key", False
    if match := _ORDER.search(text):
        order_by = match.group(1).lower()
        descending = (match.group(2) or "").lower() == "desc"
        text = text[:match.start()].strip()
        if order_by not in ORDER_FIELDS:
            raise _bad(f"Mock Jira supports ORDER BY {', '.join(ORDER_FIELDS)} only, got '{order_by}'.")
    if not text:
        raise _bad("Unbounded JQL queries are not allowed here. Please add a search restriction to your query.")
    projects: tuple[str, ...] | None = None
    conditions: list[Condition] = []
    for clause in _flatten(text):
        if m := _PROJECT_IN.fullmatch(clause):
            found = _unquote(m.group(1))
        elif m := _PROJECT_EQ.fullmatch(clause):
            found = (m.group(1).strip('"'),)
        elif m := _NAME_IN.fullmatch(clause):
            conditions.append(Condition(m.group(1).lower(), values=_unquote(m.group(2))))
            continue
        elif m := _NAME_EQ.fullmatch(clause):
            conditions.append(Condition(m.group(1).lower(), values=(m.group(2).strip('"'),)))
            continue
        elif m := _DATE.fullmatch(clause):
            conditions.append(Condition(m.group(1).lower(), op=m.group(2), bound=m.group(4)))
            continue
        else:
            raise _bad(f"Error in the JQL Query: mock Jira supports project, issuetype, status, created and updated "
                       f"restrictions joined by AND, got: {clause}")
        if projects is not None:
            raise _bad("Error in the JQL Query: mock Jira supports one project clause only.")
        projects = found
    if not projects:
        raise _bad("Unbounded JQL queries are not allowed here. Please add a search restriction to your query.")
    return Query(projects, order_by, descending, tuple(conditions))


def _bound(condition: Condition, as_of: dt.datetime) -> dt.datetime:
    if condition.bound.startswith("-"):
        amount, unit = int(condition.bound[1:-1]), condition.bound[-1]
        return as_of - dt.timedelta(days=amount * (7 if unit == "w" else 1))
    return dt.datetime.fromisoformat(condition.bound)


def matches(issue, query: Query, as_of: dt.datetime) -> bool:
    """True if the issue is in a queried project and satisfies every restriction."""
    if issue.project not in query.projects:
        return False
    for c in query.conditions:
        if c.field == "issuetype" and issue.type.casefold() not in {v.casefold() for v in c.values}:
            return False
        if c.field == "status" and issue.status.casefold() not in {v.casefold() for v in c.values}:
            return False
        if c.field in ("created", "updated"):
            value = issue.created if c.field == "created" else issue.updated
            bound = _bound(c, as_of)
            if not {">=": value >= bound, ">": value > bound, "<=": value <= bound, "<": value < bound}[c.op]:
                return False
    return True
