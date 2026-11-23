"""The workflows of Jira for the Python path (twin of modJiraConfig.ReadWorkflowRules): the projects with their issue types (project/search), then the
workflows of those issue types (workflows/preview, 25 types a call, the project administrator right). From them: the workflow of each (project, issue
type), the creation status and the moves of each workflow, the statuses that start it (first in progress from creation) and those a GLOBAL transition reaches.
It only improves what the run knows (the start of a flow, the order of a chart, the check DC28), so a call that fails gives nothing, never an error."""
import json
import logging
from urllib.parse import quote

from jira2dashboard.flow.jira_workflows import JiraWorkflows
from jira2dashboard.flow.status_rules import first_in_progress, jira_category
from jira2dashboard.standalone import jira_api as J

log = logging.getLogger("jira2dashboard.extract")
KEYS_PER_CALL, PROJECTS_PER_PAGE, TYPES_PER_PREVIEW, MAX_PAGES = 50, 50, 25, 100
REFUSED = ("J2D-JIRA-401", "J2D-JIRA-403", "J2D-JIRA-404")        # no project administrator right (modJiraConfig.IsNotAllowed): the next project may be readable
WARNING = "J2D-JIRA-030 | the workflows of Jira could not be read (%s): the start of a flow comes from the sheet Workflows, the check DC28 from the sheet Transitions"


def parse_preview(answer: dict, project: str, type_names: dict[str, str], workflow_of: dict, seen: set, initial: dict, moves: set,
                  starts: set, free: set) -> None:
    """One preview answer: (project, issue type) -> workflow name from the query contexts and, for a workflow not met yet (by its id, else its name), its
    creation status, moves, start and free statuses. A statusReference is valid only inside its answer: it is turned into a status name here."""
    names = {s.get("statusReference"): s.get("name") or "" for s in answer.get("statuses") or []}
    categories = {s.get("name"): s.get("statusCategory") for s in answer.get("statuses") or []}
    for wf in answer.get("workflows") or []:
        name = wf.get("name") or ""
        for ctx in wf.get("queryContext") or []:
            for type_id in ctx.get("issueTypes") or []:
                if str(type_id) in type_names:
                    workflow_of[(project, type_names[str(type_id)])] = name
        ident = wf.get("id") or name
        if ident in seen:
            continue                                                     # a workflow shared by several projects is read once
        seen.add(ident)
        edges, free_here = [], set()
        for t in wf.get("transitions") or []:
            target = names.get(t.get("toStatusReference"), t.get("toStatusReference") or "")
            kind = str(t.get("type") or "").upper()
            if kind == "GLOBAL":
                free_here.add(target)
            elif kind == "INITIAL":
                initial[name] = target
            elif kind == "DIRECTED":
                for link in t.get("links") or []:
                    edges.append((names.get(link.get("fromStatusReference"), link.get("fromStatusReference") or ""), target))
        moves |= {(name, source, target) for source, target in edges}
        free |= free_here
        start = first_in_progress(initial.get(name), edges, {n: c for n, c in categories.items() if jira_category(c)})
        if start:
            starts.add(start)


def read_workflows(client, keys: list[str]) -> JiraWorkflows:
    """The JiraWorkflows of the projects `keys` (empty when Jira says nothing: no right, an unreadable answer). A project whose preview Jira refuses
    (401, 403, 404) is counted and the next one is read, as the VBA does; another error stops the reading. One WARN line says either."""
    workflow_of, seen, initial, moves, starts, free, refused = {}, set(), {}, set(), set(), set(), 0
    try:
        for first in range(0, len(keys), KEYS_PER_CALL):
            start = 0
            for _page in range(MAX_PAGES):
                path = "/rest/api/3/project/search?" + "".join(f"keys={quote(k)}&" for k in keys[first:first + KEYS_PER_CALL]) \
                       + f"expand=issueTypes&startAt={start}&maxResults={PROJECTS_PER_PAGE}"
                page = client.call("GET", path, "projects")
                values = page.get("values") or []
                for project in values:
                    types = {str(t.get("id")): t.get("name") or "" for t in project.get("issueTypes") or []}
                    ids = list(types)
                    try:
                        for n in range(0, len(ids), TYPES_PER_PREVIEW):
                            body = json.dumps({"projectId": str(project.get("id")), "issueTypeIds": ids[n:n + TYPES_PER_PREVIEW]}).encode("utf-8")
                            answer = client.call("POST", "/rest/api/3/workflows/preview", "workflows", body)
                            parse_preview(answer, project.get("key") or "", types, workflow_of, seen, initial, moves, starts, free)
                    except J.JiraError as err:
                        if err.support_code not in REFUSED:
                            raise
                        refused += 1
                if page.get("isLast") is True or not values:
                    break
                start += len(values)
    except J.JiraError as err:
        log.warning(WARNING, err.support_code)
    else:
        if refused:
            log.warning(WARNING, f"{refused} project(s) refused")
    return JiraWorkflows(initial, frozenset(moves), frozenset(starts), frozenset(free), workflow_of)
