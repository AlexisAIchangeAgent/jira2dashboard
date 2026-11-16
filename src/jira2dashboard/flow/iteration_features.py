"""The Features of the ART, iteration by iteration: the intermediate level between the PI field history and Feature Predictability.

For each iteration, each PI it belongs to (the PI of the sheet PIs that holds its start) and each team (the ART as a whole, then
the team of the Feature: the value of the Team field), the Features that:
  Engaged          carry the PI at the start of the iteration and are not done yet;
  Added            carry the PI at the end of the iteration, not through the stretch that held it at the start (engaged during the
                   iteration), and are not done yet;
  Deferred         were engaged at the start, no longer carry the PI at the end, and were not released in the iteration;
  Released         are done (DoneUtc) in the iteration and carried the PI at its start or at its end;
    ReleasedPlanned   among them, those that carried the PI without a break from the baseline (the start of the first iteration of
                      the PI) to the start of the iteration: a Feature that left the PI and came back is an added one;
    ReleasedAdded     the others, engaged after the baseline;
  CarryOver        Engaged + Added - Deferred - Released: the Engaged of the next iteration of the PI (tested as a property).
The PI figures are sums of these rows (metrics.feature_predictability): the commitment is what was engaged at the baseline, so a
Feature deferred later still counts as planned, and one added later is not planned. Pure: the workbook has the same table
(IterationFeatures) in formulas, and compares with this reference.
"""
import datetime as dt
from dataclasses import dataclass

from jira2dashboard.flow.itemflow import ItemFlow
from jira2dashboard.flow.model import ART, FlowIssue, Iteration
from jira2dashboard.flow.pi_history import PIInterval, interval_at, pi_intervals


@dataclass(frozen=True)
class IterationFeatures:
    """One row of the table IterationFeatures: an iteration, a PI and a team."""
    iteration: str
    start: dt.datetime
    end: dt.datetime
    team: str                          # ART (every Feature) or the value of the Team field
    pi: str                            # the PI that holds the start of the iteration; "" when none does
    baseline: dt.datetime | None       # the start of the first iteration of that PI (where its commitment is measured)
    engaged: int
    added: int
    deferred: int
    released: int
    released_planned: int
    released_added: int
    carry_over: int


def feature_teams(issues: list[FlowIssue], feature_types: frozenset[str]) -> list[str]:
    """ART, then the teams the Features name in their Team field, alphabetical. A Feature with no team value is in the ART only."""
    named = {(i.team_value or "").strip() for i in issues if i.type in feature_types}
    return [ART] + sorted(name for name in named if name and name != ART)


def baselines(iterations: list[Iteration], pis: dict[str, tuple[dt.datetime, dt.datetime]]) -> dict[str, dt.datetime]:
    """PI -> the start of its first iteration: the moment its commitment is measured. A PI with no iteration has none."""
    out: dict[str, dt.datetime] = {}
    for pi, (start, end) in pis.items():
        starts = [i.start for i in iterations if start <= i.start < end]
        if starts:
            out[pi] = min(starts)
    return out


def _stretch(intervals: list[PIInterval], pi: str, moment: dt.datetime) -> PIInterval | None:
    """The stretch of `pi` that holds `moment`, if the Feature carries that PI then."""
    found = interval_at(intervals, moment)
    return found if found is not None and found.pi == pi else None


def _flags(intervals: list[PIInterval], done: dt.datetime | None, pi: str, iteration: Iteration, baseline: dt.datetime) -> tuple:
    """(engaged, added, deferred, released, released_added) of one Feature for one iteration and PI, as 0/1. The workbook counts
    stretches of the PI (COUNTIFS on PIIntervals): a Feature is "through" the iteration when one stretch holds its start and its
    end. Known and left: a Feature done in an iteration where it left the PI and came back counts twice in the workbook."""
    start, end = iteration.start, iteration.end
    at_start, at_end = _stretch(intervals, pi, start), _stretch(intervals, pi, end)
    through = at_start is not None and at_start is at_end
    done_before_start = done is not None and done < start
    released = done is not None and iteration.contains(done) and (at_start is not None or at_end is not None)
    engaged = at_start is not None and not done_before_start
    added = at_end is not None and not through and not done_before_start
    deferred = engaged and not through and not released
    from_baseline = at_start is not None and _stretch(intervals, pi, baseline) is at_start
    released_added = released and not (from_baseline and iteration.contains(done))
    return int(engaged), int(added), int(deferred), int(released), int(released_added)


def iteration_features(issues: list[FlowIssue], items: list[ItemFlow], iterations: list[Iteration],
                       pis: dict[str, tuple[dt.datetime, dt.datetime]], feature_types: frozenset[str]) -> list[IterationFeatures]:
    """The rows of IterationFeatures: iterations in the given order, for each the PIs that hold its start (the sheet's order; none =
    one row with an empty PI), for each the ART first and then the teams (feature_teams)."""
    done_at = {item.key: item.done for item in items}
    features = [(i, pi_intervals(i), done_at.get(i.key)) for i in issues if i.type in feature_types]
    teams = feature_teams(issues, feature_types)
    base = baselines(iterations, pis)
    out = []
    for iteration in iterations:
        holders = [pi for pi, (start, end) in pis.items() if start <= iteration.start < end] or [""]
        for pi in holders:
            totals = {team: [0] * 5 for team in teams}
            baseline = base.get(pi)
            if pi in base:
                for issue, intervals, done in features:
                    flags = _flags(intervals, done, pi, iteration, base[pi])
                    for team in {ART, (issue.team_value or "").strip()} & set(teams):
                        totals[team] = [a + b for a, b in zip(totals[team], flags)]
            for team in teams:
                engaged, added, deferred, released, released_added = totals[team]
                out.append(IterationFeatures(iteration.name, iteration.start, iteration.end, team, pi, baseline, engaged, added, deferred,
                                             released, released - released_added, released_added,
                                             engaged + added - deferred - released))
    return out
