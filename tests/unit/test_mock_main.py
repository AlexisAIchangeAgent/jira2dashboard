"""The data of `run.ps1 mock-jira`: --anomalies adds the known data-quality cases of the tests."""
from jira2dashboard.mockjira.__main__ import dataset_for, parse
from jira2dashboard.mockjira.anomalies import FIRST_NUMBER


def anomaly_keys(dataset):
    return [i.key for i in dataset.issues if int(i.key.rsplit("-", 1)[1]) >= FIRST_NUMBER]


def test_without_the_flag_the_fake_jira_serves_the_generated_issues_only():
    dataset = dataset_for(parse(["--issues", "30"]))
    assert len(dataset.issues) == 30 and anomaly_keys(dataset) == []


def test_with_the_flag_it_adds_the_known_anomalies():
    dataset = dataset_for(parse(["--issues", "30", "--anomalies"]))
    assert len(dataset.issues) > 30 and len(anomaly_keys(dataset)) == len(dataset.issues) - 30
