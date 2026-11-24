"""VelocityCheck: the Velocity of an iteration counted again from the status changes of the changelog (AL, 2026-11-18).

On the client PC the Velocity of the ART said 60 where Jira's `status CHANGED TO "Terminé" DURING` the iteration said 8. For each iteration and
team this table gives the figures that tell the causes apart, all read in the intervals of StatusIntervals (the changelog as visits of statuses):

* OtherTypes: on the ART rows, the tickets of the Program that are not Features ("ART (other types)") done in the iteration: what the ART metrics leave out;
* Velocity: the items of the team done in the iteration (DoneUtc in [Start, End), as the chart); VelocityFeatures: of which the types of
  Config.FeatureTypes (the Program project also holds other types, which Velocity counts);
* CreatedDone: items that are done in the iteration because they were CREATED in an end status (their first interval, still open): there is
  no status change in the changelog for them;
* ChangedToEnd: the status changes INTO an end status (IsEnd = 1) in the iteration, per change (what Jira's JQL counts, one more for an item that
  closes twice); ChangedToEndFeatures: of which the Feature types; ReopenedLater: those changes that the item left again (they are not its DoneUtc);
* Gap = Velocity - CreatedDone - (ChangedToEnd - ReopenedLater): 0 whenever the table explains the Velocity (a control);
* Reason (AL, 2026-11-24): why the Velocity differs from Jira's count of the changes into an end status (`status CHANGED TO ... DURING`, which
  also counts the other types of the Program), or that nothing was done; "" when the Velocity is that count.

Excel: the table VelocityCheck (formulas.VELOCITYCHECK_COLUMNS).
"""
from dataclasses import dataclass

from jira2dashboard.flow.itemflow import ItemFlow
from jira2dashboard.flow.model import ART, OTHER_TYPES, FlowIssue, Iteration, Rules


@dataclass(frozen=True)
class VelocityCheck:
    iteration: str
    team: str
    velocity: int
    velocity_features: int
    other_types: int                # the ART rows: tickets of the Program that are not Features, done in the iteration (team "ART (other types)")
    created_done: int
    changed_to_end: int
    changed_to_end_features: int
    reopened_later: int
    gap: int

    @property
    def reason(self) -> str:
        return velocity_reason(self.velocity, self.created_done, self.reopened_later, self.other_types)


NO_ITEM = "no item done in the iteration"
CREATED_DONE = " created in an end status (no status change in the changelog)"
REOPENED = " closed then reopened (a change into an end status that is not a delivery)"
OTHER_TYPES_DONE = " of other types done (left out of the ART, counted by a JQL on the Program)"


def velocity_reason(velocity: int, created_done: int, reopened_later: int, other_types: int) -> str:
    """The parts that apply, in this order, joined by "; " (Excel: TEXTJOIN of formulas.VELOCITYCHECK_COLUMNS Reason)."""
    parts = [NO_ITEM if velocity == 0 else "", f"{created_done}{CREATED_DONE}" if created_done else "",
             f"{reopened_later}{REOPENED}" if reopened_later else "", f"{other_types}{OTHER_TYPES_DONE}" if other_types else ""]
    return "; ".join(p for p in parts if p)


def velocity_check(issues: list[FlowIssue], items: list[ItemFlow], iteration: Iteration, team: str, rules: Rules,
                   feature_types: frozenset[str]) -> VelocityCheck:
    """One row: `items[n]` is the ItemFlow of `issues[n]`. Counts, so an iteration with nothing gives 0 (never not computable)."""
    velocity = velocity_features = other = created = changed = changed_features = reopened = 0
    for issue, item in zip(issues, items):
        if team == ART and item.team == OTHER_TYPES and iteration.contains(item.done):
            other += 1
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
    return VelocityCheck(iteration.name, team, velocity, velocity_features, other, created, changed, changed_features, reopened,
                         velocity - created - (changed - reopened))
