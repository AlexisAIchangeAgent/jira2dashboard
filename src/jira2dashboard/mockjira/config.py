"""The Jira configuration answers of the fake Jira: projects, issue types, statuses per issue type, workflows (preview).

Pure: no HTTP here (see server.py). Built from the dataset (which project holds which issue types, which workflow each
type follows) and from the parameter sheets Workflows and Transitions (statuses, categories, moves), so that what the
workbook proposes from these answers can be compared with the parameters. Shapes (Jira Cloud platform REST API v3, from the OpenAPI spec: documented, to verify at the client).
"""
from jira2dashboard.mockjira.dataset import Dataset
from jira2dashboard.mockjira.errors import ApiError
from jira2dashboard.params.model import ParamSheet, load_all

# Issue type -> (id, hierarchy level): 1 above the stories (Features), 0 the stories, -1 the sub-tasks.
ISSUE_TYPES = {"Feature": ("10001", 1), "Story": ("10002", 0), "Task": ("10003", 0), "Bug": ("10004", 0),
               "Sub-task": ("10005", -1)}
PROJECT_PAGE, PROJECT_PAGE_MAX = 50, 100
PREVIEW_MAX_TYPES = 25
_PREVIEW_CATEGORY = {"new": "TODO", "indeterminate": "IN_PROGRESS", "done": "DONE"}     # the preview has its own enum


def _error(status: int, message: str) -> ApiError:
    return ApiError(status, {"errorMessages": [message], "errors": {}})


def _sheets(sheets: list[ParamSheet] | None) -> dict[str, ParamSheet]:
    return {s.name: s for s in (sheets or load_all())}


def project_id(ds: Dataset, key: str) -> str:
    return str(10100 + ds.projects().index(key))


def _issue_types(ds: Dataset, key: str) -> list[str]:
    present = {i.type for i in ds.issues if i.project == key}
    return [name for name in ISSUE_TYPES if name in present]


def _workflow_of(ds: Dataset, key: str, issue_type: str) -> str:
    return next(i.workflow for i in ds.issues if i.project == key and i.type == issue_type)


def _type_json(name: str) -> dict:
    type_id, level = ISSUE_TYPES[name]
    return {"id": type_id, "name": name, "subtask": level == -1, "hierarchyLevel": level, "description": ""}


def _first(query: dict, name: str, default: str) -> str:
    return (query.get(name) or [default])[0]


def project_search(ds: Dataset, query: dict) -> dict:
    """GET /rest/api/3/project/search?keys=A&keys=B&expand=issueTypes&startAt=0&maxResults=50 (offset paging)."""
    keys = {k for value in query.get("keys", []) for k in value.split(",") if k}
    expand = {e for value in query.get("expand", []) for e in value.split(",")}
    start = int(_first(query, "startAt", "0"))
    size = min(int(_first(query, "maxResults", str(PROJECT_PAGE))), PROJECT_PAGE_MAX)
    found = [k for k in ds.projects() if not keys or k in keys]
    values = []
    for key in found[start:start + size]:
        value = {"id": project_id(ds, key), "key": key, "name": f"Project {key}", "style": "classic",
                 "simplified": False, "projectTypeKey": "software"}
        if "issueTypes" in expand:
            value["issueTypes"] = [_type_json(name) for name in _issue_types(ds, key)]
        values.append(value)
    return {"startAt": start, "maxResults": size, "total": len(found), "isLast": start + size >= len(found),
            "values": values}


def _statuses(workflow: str, sheets: dict[str, ParamSheet]) -> list[str]:
    return [row[2] for row in sheets["Workflows"].rows if row[0] == workflow]


def project_statuses(ds: Dataset, key: str, sheets: list[ParamSheet] | None = None) -> list[dict]:
    """GET /rest/api/3/project/{key}/statuses: the statuses of each issue type, with the Jira status category."""
    if key not in ds.projects():
        raise _error(404, f"No project could be found with key '{key}'.")
    params = _sheets(sheets)
    answer = []
    for name in _issue_types(ds, key):
        workflow = _workflow_of(ds, key, name)
        statuses = []
        for status in _statuses(workflow, params):
            category_key, category_name = ds.catalog.jira_category(workflow, status)
            statuses.append({"id": ds.catalog.ids[status], "name": status, "description": "",
                             "statusCategory": {"key": category_key, "name": category_name}})
        answer.append({"id": ISSUE_TYPES[name][0], "name": name, "subtask": name == "Sub-task", "statuses": statuses})
    return answer


def workflows_preview(ds: Dataset, body: dict, sheets: list[ParamSheet] | None = None) -> dict:
    """POST /rest/api/3/workflows/preview {"projectId": "...", "issueTypeIds": [...]}: the workflows of these issue
    types in the project, with their statuses and transitions. A statusReference is only valid inside this answer."""
    if not body.get("projectId"):
        raise _error(400, "projectId is required.")
    type_ids = [str(t) for t in body.get("issueTypeIds") or []]
    if not type_ids:
        raise _error(400, "Provide at least one of workflowNames, workflowIds or issueTypeIds.")
    if len(type_ids) > PREVIEW_MAX_TYPES:
        raise _error(400, f"At most {PREVIEW_MAX_TYPES} issue type IDs can be specified.")
    key = next((k for k in ds.projects() if project_id(ds, k) == str(body["projectId"])), None)
    if key is None:
        raise _error(404, "The previews were not found.")
    params = _sheets(sheets)
    by_workflow: dict[str, list[str]] = {}
    for name in _issue_types(ds, key):
        if ISSUE_TYPES[name][0] in type_ids:
            by_workflow.setdefault(_workflow_of(ds, key, name), []).append(ISSUE_TYPES[name][0])
    refs: dict[str, str] = {}
    statuses: list[dict] = []

    def ref(workflow: str, status: str) -> str:
        if status not in refs:
            refs[status] = f"ref-{len(refs) + 1:03d}"
            category = _PREVIEW_CATEGORY[ds.catalog.jira_category(workflow, status)[0]]
            statuses.append({"id": ds.catalog.ids[status], "name": status, "rawName": status, "description": "",
                             "statusCategory": category, "statusReference": refs[status], "scope": {"type": "GLOBAL"}})
        return refs[status]

    workflows = [_workflow_json(ds, workflow, by_workflow[workflow], project_id(ds, key), params, ref)
                 for workflow in by_workflow]
    return {"statuses": statuses, "workflows": workflows}


def _workflow_json(ds: Dataset, workflow: str, type_ids: list[str], pid: str, params: dict, ref) -> dict:
    rows = [row for row in params["Workflows"].rows if row[0] == workflow]
    status_refs = [{"statusReference": ref(workflow, row[2]), "deprecated": False} for row in rows]
    transitions = [{"id": "1", "name": "Create", "type": "INITIAL", "links": [],
                    "toStatusReference": ref(workflow, next(row[2] for row in rows if row[1] == 1))}]
    targets: dict[str, list[str]] = {}                  # one DIRECTED transition per target, one link per source
    for wf, source, target, _kind in params["Transitions"].rows:
        if wf == workflow:
            targets.setdefault(target, []).append(source)
    for target, sources in targets.items():
        transitions.append({"id": str(10 * len(transitions) + 1), "name": target, "type": "DIRECTED",
                            "toStatusReference": ref(workflow, target),
                            "links": [{"fromStatusReference": ref(workflow, s), "fromPort": 0, "toPort": 1}
                                      for s in sources]})
    for row in rows:
        if row[6] == 1:                                 # a free status is reachable from any status: a GLOBAL transition
            transitions.append({"id": str(10 * len(transitions) + 1), "name": row[2], "type": "GLOBAL", "links": [],
                                "toStatusReference": ref(workflow, row[2])})
    return {"id": f"wf-{workflow.lower()}", "name": workflow, "description": "", "scope": {"type": "GLOBAL"},
            "queryContext": [{"project": pid, "issueTypes": type_ids}], "statuses": status_refs,
            "transitions": transitions}
