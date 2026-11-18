"""What Jira says about workflows (AL, 2026-11-18): the workflow of each issue type of each project, its creation status, its moves, the statuses that start
it and those reached from anywhere. The sheets JiraIssueTypes, JiraWorkflows and JiraTransitions hold it in the workbook; the Python path reads it from the
same calls (standalone/workflows.py) and the fake Jira's definition gives it to the tests (mockjira/flow_adapter.jira_workflows). Pure.

It feeds: the start and the free statuses of the status rules (status_rules), the order of the stack of the chart Cycle time by status (feature_depths),
and the check DC28 (a move that the workflow of the ticket does not allow).
"""
from dataclasses import dataclass, field
from typing import Iterable

from jira2dashboard.flow.status_rules import flow_depths


@dataclass(frozen=True)
class JiraWorkflows:
    initial: dict[str, str] = field(default_factory=dict)                       # workflow -> its creation status
    moves: frozenset[tuple[str, str, str]] = frozenset()                       # (workflow, from status, to status)
    starts: frozenset[str] = frozenset()                                       # the statuses that start a workflow (first in progress from creation)
    free: frozenset[str] = frozenset()                                         # the statuses a GLOBAL transition reaches
    workflow_of: dict[tuple[str, str], str] = field(default_factory=dict)      # (project, issue type) -> workflow name

    def knows(self, workflow: str) -> bool:
        """True when the moves of this workflow were read: only then can a move be judged."""
        return any(wf == workflow for wf, _source, _target in self.moves)

    def allowed(self, workflow: str, source: str, target: str) -> bool:
        return (workflow, source, target) in self.moves

    def feature_depths(self, programs: Iterable[str], feature_types: Iterable[str]) -> dict[str, int]:
        """status -> its smallest depth (moves from the creation status) in the workflows the Feature types of the Program projects follow."""
        names = {self.workflow_of[(project, kind)] for project in programs for kind in feature_types if (project, kind) in self.workflow_of}
        depths: dict[str, int] = {}
        for name in sorted(names):
            moves = [(source, target) for wf, source, target in self.moves if wf == name]
            for status, depth in flow_depths(self.initial.get(name), moves).items():
                depths[status] = min(depth, depths.get(status, depth))
        return depths
