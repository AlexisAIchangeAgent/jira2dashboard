"""The Example column written next to Iteration-label | Start | End (sheets Iterations and PIs): pure checks, no Excel."""
import datetime as dt
import re

from jira2dashboard.params.write import EXAMPLES


def test_iterations_and_pis_have_one_example_per_column():
    assert set(EXAMPLES) == {"Iterations", "PIs"}
    assert all(len(lines) == 3 for lines in EXAMPLES.values())
    for lines, labels in ((EXAMPLES["Iterations"], ("Iteration-label:", "Start:", "End:")), (EXAMPLES["PIs"], ("PI:", "Start:", "End:"))):
        assert all(line.startswith(label) for line, label in zip(lines, labels))


def test_the_example_dates_are_real_dates_and_the_end_follows_the_start():
    for lines in EXAMPLES.values():
        start, end = (dt.date.fromisoformat(re.search(r"\d{4}-\d{2}-\d{2}", line).group()) for line in lines[1:])
        assert start < end
