"""VelocityCheck: the Velocity of an iteration counted again from the status changes of the changelog (AL, 2026-11-18).

On the client PC the Velocity of the ART said 60 where Jira's `status CHANGED TO "Terminé" DURING` the iteration said 8. For each iteration and
team this table gives the figures that tell the causes apart, all read in the intervals of StatusIntervals (the changelog as visits of statuses):

* Velocity: the items of the team done in the iteration (DoneUtc in [Start, End), as the chart); VelocityFeatures: of which the types of
  Config.FeatureTypes (the Program project also holds other types, which Velocity counts);
* CreatedDone: items that are done in the iteration because they were CREATED in an end status (their first interval, still open): there is
  no status change in the changelog for them;
* ChangedToEnd: the status changes INTO an end status (IsEnd = 1) in the iteration, per change (what Jira's JQL counts, one more for an item that
  closes twice); ChangedToEndFeatures: of which the Feature types; ReopenedLater: those changes that the item left again (they are not its DoneUtc);
* Gap = Velocity - CreatedDone - (ChangedToEnd - ReopenedLater): 0 whenever the table explains the Velocity (a control).

Excel: the table VelocityCheck (formulas.VELOCITYCHECK_COLUMNS).
"""
from dataclasses import dataclass

from jira2dashboard.flow.itemflow import ItemFlow
from jira2dashboard.flow.model import FlowIssue, Iteration, Rules


@dataclass(frozen=True)
class VelocityCheck:
    iteration: str
    team: str
    velocity: int
    velocity_features: int
    created_done: int
    changed_to_end: int
    changed_to_end_features: int
    reopened_later: int
    gap: int


def velocity_check(issues: list[FlowIssue], items: list[ItemFlow], iteration: Iteration, team: str, rules: Rules,
                   feature_types: frozenset[str]) -> VelocityCheck:
    """One row: `items[n]` is the ItemFlow of `issues[n]`. Counts, so an iteration with nothing gives 0 (never not computable)."""
    velocity = velocity_features = created = changed = changed_features = reopened = 0
    for issue, item in zip(issues, items):
        if item.team != team:
            continue
        feature = issue.type in feature_types
        if iteration.contains(item.done):
            velocity += 1
            velocity_features += feature
        for visit in item.intervals:
            if not rules.rule(issue.workflow, visit.status).is_end or not iteration.contains(visit.entered):
                continue
            if visit.seq == 1:
                created += visit.is_open                         # created in an end status and still there: done at its creation
            else:
                changed += 1
                changed_features += feature
                reopened += not visit.is_open
    return VelocityCheck(iteration.name, team, velocity, velocity_features, created, changed, changed_features, reopened,
                         velocity - created - (changed - reopened))
