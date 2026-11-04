"""The fake Jira's issues as flow-metric inputs: what the extraction plus the sheets Projects and Workflows give."""
import datetime as dt

from jira2dashboard.flow.model import ART, Change, FlowIssue
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


def to_flow_issues(dataset: Dataset, sheets: list[ParamSheet] | None = None) -> list[FlowIssue]:
    """Program issues are reported under the team ART and follow the Feature workflow; every other project
    follows the Team workflow and reports under its team name (sheet Projects)."""
    projects = {row[0]: (row[1], row[2]) for s in (sheets or load_all()) if s.name == "Projects" for row in s.rows}
    out = []
    for issue in dataset.issues:
        role, team = projects[issue.project]
        program = role == "Program"
        workflow = "Feature" if program else "Team"
        resolution = dataset.resolution(issue)
        sprint = current_sprint(issue)
        out.append(FlowIssue(
            key=issue.key, team=ART if program else team, type=issue.type, work_type=issue.work_type,
            workflow=workflow, created=issue.created, status=issue.status,
            changes=tuple(Change(c.id, c.at, c.from_status, c.to_status) for c in issue.changes),
            parent=issue.parent, resolved=resolution[1] if resolution else None,
            status_category=dataset.catalog.jira_category(workflow, issue.status)[0],
            fix_versions=issue.fix_versions, story_points=issue.story_points, pi=issue.pi,
            business_value=issue.business_value, team_value=issue.team,
            sprint_state=sprint.state if sprint else "", sprint_start=sprint.start if sprint else None,
            sprint_end=sprint.end if sprint else None, sprint_count=len(issue.sprints),
            target_start=_day(issue.target_start), target_end=_day(issue.target_end), blocks=issue.blocks,
            key_changes=len(issue.key_changes)))
    return out
