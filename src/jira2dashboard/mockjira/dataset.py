"""Domain of the fake Jira: issues with their status history, and the catalogue of statuses and fields.

All datetimes are naive UTC. The server formats them with a fixed offset (like Jira, in the user's time zone).
"""
import datetime as dt
from dataclasses import dataclass, field

from jira2dashboard.params.model import ParamSheet, load_all

# Jira status categories (name shown in the French UI) for each category of our Workflows sheet.
_JIRA_CATEGORY = {
    "Backlog": ("new", "À faire"),
    "Active": ("indeterminate", "En cours"),
    "Wait": ("indeterminate", "En cours"),
    "Done": ("done", "Terminé"),
    "Removed": ("done", "Terminé"),
    "Unclassified": ("new", "À faire"),
}


@dataclass(frozen=True)
class FieldIds:
    """Ids of the custom fields in this fake Jira (mock choices: at the client they are read from GET /field)."""
    story_points: str = "customfield_10016"
    work_type: str = "customfield_10200"
    pi: str = "customfield_10300"
    business_value: str = "customfield_10301"

    def catalog(self) -> list[dict]:
        """GET /rest/api/3/field: system and custom fields (id, name, custom, schema)."""
        def system(fid, name, kind):
            return {"id": fid, "name": name, "custom": False, "schema": {"type": kind, "system": fid}}

        def custom(fid, name, kind):
            return {"id": fid, "name": name, "custom": True, "schema": {"type": kind, "customId": int(fid.split("_")[1])}}
        return [
            system("summary", "Summary", "string"), system("status", "Status", "status"),
            system("issuetype", "Issue Type", "issuetype"), system("project", "Project", "project"),
            system("created", "Created", "datetime"), system("updated", "Updated", "datetime"),
            system("resolution", "Resolution", "resolution"), system("resolutiondate", "Resolved", "datetime"),
            system("duedate", "Due date", "date"), system("labels", "Labels", "array"),
            system("fixVersions", "Fix versions", "array"), system("parent", "Parent", "issuelink"),
            system("priority", "Priority", "priority"), system("assignee", "Assignee", "user"),
            custom(self.story_points, "Story point estimate", "number"),
            custom(self.work_type, "Work type", "string"),
            custom(self.pi, "PI", "string"),
            custom(self.business_value, "Business Value", "number"),
        ]


@dataclass(frozen=True)
class StatusCatalog:
    """Status ids and categories, built from the workflow parameters (params/workflows.md)."""
    ids: dict[str, str]                                   # status name -> id
    categories: dict[tuple[str, str], str]                # (workflow, status) -> Active/Wait/Done/Removed/Backlog
    end_statuses: dict[str, str]                          # workflow -> status with IsEnd = 1

    @classmethod
    def from_params(cls, sheets: list[ParamSheet] | None = None) -> "StatusCatalog":
        workflows = next(s for s in (sheets or load_all()) if s.name == "Workflows")
        names = sorted({row[2] for row in workflows.rows})
        return cls(
            ids={name: str(10000 + i) for i, name in enumerate(names, start=1)},
            categories={(row[0], row[2]): row[3] for row in workflows.rows},
            end_statuses={row[0]: row[2] for row in workflows.rows if row[5] == 1},
        )

    def jira_category(self, workflow: str, status: str) -> tuple[str, str]:
        """(key, name) of the Jira status category: new / indeterminate / done."""
        return _JIRA_CATEGORY[self.categories.get((workflow, status), "Unclassified")]


@dataclass(frozen=True)
class Change:
    """One status change in an issue history."""
    id: str
    at: dt.datetime
    from_status: str
    to_status: str
    author: str = "Ann Martin"
    author_id: str = "acc-ann"


@dataclass(frozen=True)
class Issue:
    key: str
    id: str
    project: str
    type: str                          # Feature, Story, Task, Bug, Sub-task
    work_type: str                     # TypeCategorie: Feature, Enabler, Autre, Defect
    summary: str
    created: dt.datetime
    changes: tuple[Change, ...] = ()
    initial_status: str = "Nouveau"
    workflow: str = "Team"             # "Feature" for features, "Team" for everything else
    parent: str | None = None
    priority: str = "Medium"
    assignee: str | None = None
    labels: tuple[str, ...] = ()
    fix_versions: tuple[str, ...] = ()
    story_points: float | None = None
    business_value: int | None = None
    pi: str | None = None
    due: dt.date | None = None
    resolution_override: tuple[dt.datetime | None] | None = None    # anomaly: (resolved at,) or (None,) = no resolution date

    @property
    def status(self) -> str:
        return self.changes[-1].to_status if self.changes else self.initial_status

    @property
    def updated(self) -> dt.datetime:
        return self.changes[-1].at if self.changes else self.created

    @property
    def number(self) -> int:
        return int(self.key.rsplit("-", 1)[1])


@dataclass(frozen=True)
class Dataset:
    issues: tuple[Issue, ...]
    as_of: dt.datetime
    catalog: StatusCatalog = field(default_factory=StatusCatalog.from_params)
    fields: FieldIds = field(default_factory=FieldIds)

    def __post_init__(self):
        keys = [i.key for i in self.issues]
        ids = [i.id for i in self.issues]
        if len(set(keys)) != len(keys) or len(set(ids)) != len(ids):
            raise ValueError("duplicate issue key or id")

    def by_key_or_id(self, ref: str) -> Issue | None:
        for issue in self.issues:
            if ref in (issue.key, issue.id):
                return issue
        return None

    def projects(self) -> list[str]:
        return sorted({i.project for i in self.issues})

    def resolution(self, issue: Issue) -> tuple[str, dt.datetime] | None:
        """(resolution name, resolved at) when the issue sits in an end or removed status; else None."""
        if issue.resolution_override is not None:
            at = issue.resolution_override[0]
            return None if at is None else ("Done", at)
        status = issue.status
        category = self.catalog.categories.get((issue.workflow, status))
        if category not in ("Done", "Removed"):
            return None
        return ("Done" if category == "Done" else "Won't Do", issue.updated)


def issue(key: str, *, moves: tuple[tuple[str, str], ...] = (), created: str = "2026-10-01T09:00",
          type: str = "Story", project: str | None = None, **fields) -> Issue:
    """Hand-written case: moves = (("2026-10-02T10:00", "En analyse"), ...), each a move to that status.

    The id is 10000 + the key number: pass id=... when several projects share the same numbers.
    """
    number = int(key.rsplit("-", 1)[1])
    changes = []
    previous = fields.pop("initial_status", "Nouveau")
    initial = previous
    for n, (at, status) in enumerate(moves, start=1):
        changes.append(Change(id=f"{number}{n:03d}", at=dt.datetime.fromisoformat(at), from_status=previous,
                              to_status=status))
        previous = status
    return Issue(
        key=key, id=fields.pop("id", str(10000 + number)), project=project or key.rsplit("-", 1)[0], type=type,
        work_type=fields.pop("work_type", "Autre"), summary=fields.pop("summary", f"Summary of {key}"),
        created=dt.datetime.fromisoformat(created), changes=tuple(changes), initial_status=initial,
        workflow=fields.pop("workflow", "Feature" if type == "Feature" else "Team"), **fields)
