"""The named scenarios of the fake Jira (`mock-jira --scenario NAME`): a small, known history instead of generated data."""
from jira2dashboard.mockjira import art_metrics
from jira2dashboard.mockjira.dataset import Scenario

SCENARIOS = {art_metrics.NAME: art_metrics.scenario}


def get(name: str) -> Scenario:
    """The scenario `name`; a ValueError that lists the known names otherwise."""
    if name not in SCENARIOS:
        raise ValueError(f"unknown scenario {name!r}: known scenarios are {', '.join(SCENARIOS)}")
    return SCENARIOS[name]()
