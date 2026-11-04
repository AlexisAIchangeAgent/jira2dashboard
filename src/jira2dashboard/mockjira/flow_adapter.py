"""The fake Jira's issues as flow-metric inputs: what the extraction plus the sheets Projects and Workflows give."""
from jira2dashboard.flow.model import ART, Change, FlowIssue
from jira2dashboard.mockjira.dataset import Dataset
from jira2dashboard.params.model import ParamSheet, load_all


def to_flow_issues(dataset: Dataset, sheets: list[ParamSheet] | None = None) -> list[FlowIssue]:
    """Program issues are reported under the team ART and follow the Feature workflow; every other project
    follows the Team workflow and reports under its team name (sheet Projects)."""
    projects = {row[0]: (row[1], row[2]) for s in (sheets or load_all()) if s.name == "Projects" for row in s.rows}
    out = []
    for issue in dataset.issues:
        role, team = projects[issue.project]
        program = role == "Program"
        out.append(FlowIssue(
            key=issue.key, team=ART if program else team, type=issue.type, work_type=issue.work_type,
            workflow="Feature" if program else "Team", created=issue.created, status=issue.status,
            changes=tuple(Change(c.id, c.at, c.from_status, c.to_status) for c in issue.changes)))
    return out
