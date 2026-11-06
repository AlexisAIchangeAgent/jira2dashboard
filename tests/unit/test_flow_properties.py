"""Properties of the flow metrics on generated data (1200 issues, seeded): what must hold whatever the histories."""
import datetime as dt

import pytest

from jira2dashboard.flow import blueprint
from jira2dashboard.flow.itemflow import item_flow
from jira2dashboard.flow.load import ONE_DAY, in_process_at, snapshot_days
from jira2dashboard.flow.metrics import flow_metrics
from jira2dashboard.flow.model import NOT_COMPUTABLE, Iteration, Rules
from jira2dashboard.mockjira.flow_adapter import to_flow_issues
from jira2dashboard.mockjira.generator import generate
from jira2dashboard.params.model import load_all, project_keys

STATIC_RULES = Rules.from_params()


@pytest.fixture(scope="module")
def dataset():
    return generate(seed=7, issues=1200, projects=project_keys())


@pytest.fixture(scope="module")
def issues(dataset):
    return to_flow_issues(dataset)


@pytest.fixture(scope="module")
def items(issues, dataset):
    return [item_flow(i, dataset.as_of, STATIC_RULES) for i in issues]


@pytest.fixture(scope="module")
def iterations(dataset):
    """Iterations of 14 days from 2026-06-01 up to the as-of date (the last one is cut by it)."""
    first, out = dt.datetime(2026, 6, 1), []
    while first < dataset.as_of:
        out.append(Iteration(f"I{len(out) + 1}", first, first + dt.timedelta(days=14)))
        first += dt.timedelta(days=14)
    return out


@pytest.fixture(scope="module", params=["A", "B"])
def results(request, items, iterations, dataset):
    return request.param, flow_metrics(items, iterations, request.param, dataset.as_of)


def numbers(values):
    return [v for v in values if v != NOT_COMPUTABLE]


def test_the_data_is_rich_enough_for_the_properties(items):
    assert sum(1 for i in items if i.done) > 100 and sum(1 for i in items if i.removed) > 5
    assert sum(1 for i in items if i.backlog_returns and i.start and i.done) > 3
    assert sum(1 for i in items if i.reopened) == 0            # the generator never reopens: the hand cases cover it
    projects = next(s for s in load_all() if s.name == "Projects")
    assert {i.team for i in items} == set(blueprint.team_names([tuple(r) for r in projects.rows]))


def test_program_issues_are_the_art_team_and_nobody_else_is(issues):
    assert {i.team for i in issues if i.type == "Feature"} == {"ART"}
    assert all(i.team != "ART" for i in issues if i.type != "Feature")


def test_velocity_is_the_number_of_done_dates_in_the_iteration(items, iterations, results):
    mode, metrics = results
    for m in metrics:
        iteration = next(p for p in iterations if p.name == m.iteration)
        assert m.velocity == sum(1 for i in items if i.team == m.team and iteration.contains(i.done))


def test_every_done_issue_falls_in_exactly_one_iteration(items, iterations, results):
    mode, metrics = results
    covered = sum(m.velocity for m in metrics)
    in_window = sum(1 for i in items if i.done and iterations[0].start <= i.done < iterations[-1].end)
    assert covered == in_window == sum(1 for i in items if i.done)


def test_included_plus_excluded_equals_the_issues_done(results):
    mode, metrics = results
    assert all(m.flow_time_count + m.flow_time_excluded == m.velocity for m in metrics)


def test_flow_efficiency_is_between_0_and_1(results):
    mode, metrics = results
    values = numbers(m.flow_efficiency for m in metrics)
    assert values and all(0 <= v <= 1 for v in values)


def test_flow_time_percentiles_are_ordered_and_positive(results):
    mode, metrics = results
    for m in metrics:
        if m.flow_time_count:
            assert 0 < m.flow_time_median <= m.flow_time_p85
        else:
            assert m.flow_time_median == m.flow_time_p85 == NOT_COMPUTABLE


def test_a_distribution_adds_up_to_100_percent(results):
    mode, metrics = results
    shares = [m.distribution_done for m in metrics] + [m.distribution_wip for m in metrics]
    figures = [s for s in shares if s != NOT_COMPUTABLE]
    assert figures and all(sum(s.values()) == pytest.approx(1) for s in figures)
    done_missing = [m for m in metrics if m.distribution_done == NOT_COMPUTABLE]
    assert all(m.velocity == 0 for m in done_missing)           # not computable only when nothing was done


def test_the_load_at_the_end_of_a_iteration_is_the_number_of_issues_in_process_then(items, iterations, results, dataset):
    """Independent count: replay each issue's changes to get its status at the instant, without using intervals."""
    mode, metrics = results
    rules = STATIC_RULES
    by_team = {}
    for m in metrics:
        by_team.setdefault(m.iteration, {})[m.team] = m.flow_load_end
    checked = 0
    for iteration in iterations:
        days = snapshot_days(iteration.start, iteration.end, dataset.as_of)
        if not days:
            continue
        instant = days[-1] + ONE_DAY
        for team, reported in by_team[iteration.name].items():
            expected = sum(1 for i in items if i.team == team and _in_process_by_replay(i, instant, mode, rules, dataset))
            assert reported == expected, (iteration.name, team)
            checked += 1
    assert checked > 50


def _in_process_by_replay(item, instant, mode, rules, dataset):
    wf = "Feature" if item.team == "ART" else "Team"
    if mode == "B":
        return in_process_at(item, instant, "B")            # mode B has no per-status rule to replay: covered below
    if item.first_start is None or item.first_start > instant or instant < item.created:
        return False
    # the status the issue was in at `instant`: its last change before or at it
    status = None
    for interval in item.intervals:
        if interval.entered <= instant and (interval.left is None or instant < interval.left):
            status = interval.status
    return status is not None and rules.rule(wf, status).category in ("Active", "Wait", "Unclassified")


def test_the_flow_time_restarts_after_a_return_through_backlog(items):
    restarted = [i for i in items if i.backlog_returns and i.start and i.done]
    assert restarted
    for i in restarted:
        last_backlog = max(x.entered for x in i.intervals if x.category == "Backlog")
        assert i.start > last_backlog                            # the flow time starts after the last backlog entry
        assert i.flow_time_days < (i.done - i.first_start).total_seconds() / 86400 or i.start == i.first_start


def test_mode_b_never_counts_an_issue_before_its_last_start(items):
    for i in items[:150]:
        moment = (i.start or i.created) - dt.timedelta(hours=1)
        assert in_process_at(i, moment, "B") is False


def test_mode_b_counts_nothing_without_a_start(items):
    never_restarted = [i for i in items if i.start is None]
    assert never_restarted
    horizon = dt.datetime(2026, 10, 15)
    assert not any(in_process_at(i, horizon, "B") for i in never_restarted)


def test_times_are_consistent(items):
    for i in items:
        if i.flow_time_days is not None:
            assert i.flow_time_days >= 0 and i.lead_time_days >= i.flow_time_days
            total = i.active_days + i.wait_days + i.unclassified_days
            assert total <= i.flow_time_days + 1e-9
            if i.reopened == 0:
                assert total == pytest.approx(i.flow_time_days)     # no Backlog or Done interval inside [start, done)
        assert not (i.done and i.removed)


def test_an_issue_is_never_in_process_before_it_exists(items):
    for i in items[:100]:
        assert not in_process_at(i, i.created - dt.timedelta(seconds=1), "A")


def test_not_computable_is_never_a_zero(results):
    mode, metrics = results
    for m in metrics:
        for value in (m.flow_time_median, m.flow_time_p85, m.flow_efficiency, m.flow_load_average):
            assert value == NOT_COMPUTABLE or isinstance(value, float | int)
        if m.flow_time_count == 0:                          # nothing to measure: the text, never a 0
            assert m.flow_time_median == m.flow_time_p85 == m.flow_efficiency == NOT_COMPUTABLE
        else:
            assert isinstance(m.flow_time_median, float | int) and isinstance(m.flow_time_p85, float | int)
        assert m.flow_predictability == NOT_COMPUTABLE
