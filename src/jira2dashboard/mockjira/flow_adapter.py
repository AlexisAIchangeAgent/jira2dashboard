"""The fake Jira's issues as flow-metric inputs: what the extraction plus the sheets Projects and Workflows give."""
import datetime as dt

from jira2dashboard.flow.checks import CheckSettings
from jira2dashboard.flow.jira_workflows import JiraWorkflows
from jira2dashboard.flow.model import Change, FlowIssue, PIChange, Rules, program_team
from jira2dashboard.flow.status_rules import first_in_progress
from jira2dashboard.mockjira.dataset import Dataset, Issue
from jira2dashboard.params.model import ParamSheet, load_all


def current_sprint(issue: Issue):
    """The sprint that counts for the checks: the active one, else a future one, else the last closed one."""
    for state in ("active", "future", "closed"):
        found = [s for s in issue.sprints if s.state == state]
        if found:
            return found[-1]
    return None


def _day(value: dt.date | None) -> dt.datetime | None:
    return dt.datetime(value.year, value.month, value.day) if value else None


def to_flow_issues(dataset: Dataset, sheets: list[ParamSheet] | None = None,
                   feature_types: frozenset[str] | None = None) -> list[FlowIssue]:
    """Program issues follow the Feature workflow and are reported under the team ART when their type is in `feature_types` (None: the parameter
    FeatureTypes of the sheet Config, as the workbook does), under "ART (other types)" otherwise; every other project follows the Team workflow
    and reports under its team name (sheet Projects)."""
    used = list(sheets or dataset.sheets or load_all())
    if feature_types is None:
        feature_types = CheckSettings.from_params(used).feature_types
    projects = {row[0]: (row[1], row[2] or row[0]) for s in used if s.name == "Projects" for row in s.rows}
    out = []
    for issue in dataset.issues:
        role, team = projects[issue.project]
        program = role == "Program"
        workflow = "Feature" if program else "Team"
        resolution = dataset.resolution(issue)
        sprint = current_sprint(issue)
        out.append(FlowIssue(
            key=issue.key, team=program_team(issue.type, feature_types) if program else team, type=issue.type, work_type=issue.work_type,
            workflow=workflow, created=issue.created, status=issue.status,
            changes=tuple(Change(c.id, c.at, c.from_status, c.to_status) for c in issue.changes),
            parent=issue.parent, resolved=resolution[1] if resolution else None,
            status_category=dataset.catalog.jira_category(workflow, issue.status)[0],
            fix_versions=issue.fix_versions, story_points=issue.story_points, pi=issue.pi,
            business_value=issue.business_value, team_value=issue.team,
            sprint_state=sprint.state if sprint else "", sprint_start=sprint.start if sprint else None,
            sprint_end=sprint.end if sprint else None, sprint_count=len(issue.sprints),
            target_start=_day(issue.target_start), target_end=_day(issue.target_end), blocks=issue.blocks,
            key_changes=len(issue.key_changes), jira_workflow=dataset.workflow_names.get(issue.workflow, issue.workflow),
            pi_changes=tuple(PIChange(c.id, c.at, c.from_string, c.to_string) for c in issue.field_changes
                             if c.field_id == dataset.fields.pi)))
    return out


def jira_workflows(dataset: Dataset) -> JiraWorkflows:
    """What the fake Jira says about the workflows of this dataset (its definition in full: the sheets of `jira_sheets`): what the sheets JiraIssueTypes,
    JiraWorkflows and JiraTransitions hold after the workbook reads the configuration."""
    jira = {sheet.name: sheet for sheet in (dataset.jira_sheets or dataset.sheets or load_all())}
    categories = {}                                              # status -> the key of its Jira category (the first workflow that has it)
    for (workflow, status), _cat in dataset.catalog.categories.items():
        categories.setdefault(status, dataset.catalog.jira_category(workflow, status)[0])
    initial, starts, free, moves = {}, set(), set(), set()
    for kind in sorted({row[0] for row in jira["Workflows"].rows}):
        rows = [row for row in jira["Workflows"].rows if row[0] == kind]
        name = dataset.workflow_names.get(kind, kind)
        edges = [(source, target) for wf, source, target in jira["Transitions"].rows if wf == kind]
        initial[name] = rows[0][1]                               # the first row is the status of creation
        moves |= {(name, source, target) for source, target in edges}
        start = first_in_progress(rows[0][1], edges, categories)
        if start:
            starts.add(start)
        free |= {row[1] for row in rows if row[5] == 1}
    projects = {i.project for i in dataset.issues}
    workflow_of = {(i.project, i.type): dataset.workflow_names.get(i.workflow, i.workflow) for i in dataset.issues if i.project in projects}
    return JiraWorkflows(initial, frozenset(moves), frozenset(starts), frozenset(free), workflow_of)


def jira_rules(dataset: Dataset) -> Rules:
    """The status rules the workbook ends up with for this dataset: what the fake Jira defines (the catalog of the statuses, the category of the
    current status of each ticket, the start of each workflow by its transitions, the free statuses) under the exceptions of the sheet Workflows."""
    own = list(dataset.sheets or load_all())
    categories = {}
    for (workflow, status), _cat in dataset.catalog.categories.items():
        categories.setdefault(status, dataset.catalog.jira_category(workflow, status)[0])
    current = {i.status: categories[i.status] for i in dataset.issues if i.status in categories}
    info = jira_workflows(dataset)
    return Rules.from_params(own, categories, current, set(info.starts), set(info.free))
