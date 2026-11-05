"""Prepare a copy of a workbook to try it by hand against the fake Jira (scripts\\manual-test.ps1 runs it).

The copy, out\\inter\\jira2dashboard_mock.xlsm, gets what a client would type: in Config the address, account,
reference date and custom field ids of the fake Jira of `run.ps1 mock-jira`; 14-day iterations in Periods; the PI of
the known anomalies in PIs. The workbook given is never changed: the one to deliver keeps its client values empty.
"""
import datetime as dt
import shutil
import sys
from pathlib import Path

from jira2dashboard.cli.args import parser
from jira2dashboard.excel import com
from jira2dashboard.flow.model import Period
from jira2dashboard.mockjira.anomalies import PI, PI_WINDOW
from jira2dashboard.mockjira.dataset import FieldIds
from jira2dashboard.mockjira.generator import AS_OF, SINCE
from jira2dashboard.mockjira.server import MOCK_EMAIL, MOCK_TOKEN
from jira2dashboard.paths import INTER_DIR, RELEASE_DIR

COPY_NAME = "jira2dashboard_mock.xlsm"
COPY = INTER_DIR / COPY_NAME
FIELD_IDS = {"StoryPointsField": "story_points", "BusinessValueField": "business_value", "PIField": "pi",
             "TeamField": "team", "SprintField": "sprint", "TargetStartField": "target_start",
             "TargetEndField": "target_end"}          # Config parameter of a data check -> attribute of FieldIds
PIS = [Period(PI, *PI_WINDOW)]
ROWS = "A2:C200"                                      # the rows of Periods and PIs that the copy rewrites


def base_url(port: int) -> str:
    return f"http://127.0.0.1:{port}"


def config_values(url: str, fields: FieldIds = FieldIds(), as_of: dt.datetime = AS_OF) -> dict[str, str]:
    """Parameter -> value of the Config sheet, for the fake Jira at `url`."""
    return {"BaseUrl": url, "Email": MOCK_EMAIL, "ApiToken": MOCK_TOKEN, "AsOf": f"{as_of:%Y-%m-%d %H:%M}",
            "WorkTypeField": fields.work_type} | {name: getattr(fields, attr) for name, attr in FIELD_IDS.items()}


def iterations(first: dt.datetime = SINCE, as_of: dt.datetime = AS_OF, days: int = 14) -> list[Period]:
    """I1, I2... of `days` days from `first`, one after the other, the last one holding `as_of`."""
    periods, start = [], first
    while start < as_of:
        end = start + dt.timedelta(days=days)
        periods.append(Period(f"I{len(periods) + 1}", start, end))
        start = end
    return periods


def write_periods(ws, periods: list[Period]) -> None:
    """Rows Name | Start | End of the sheet Periods or PIs; the rows of an earlier run are cleared first."""
    ws.Range(ROWS).ClearContents()
    for row, period in enumerate(periods, start=2):
        ws.Cells(row, 1).Value = com.as_cell(period.name)
        ws.Cells(row, 2).Value2 = com.to_excel_serial(period.start)
        ws.Cells(row, 3).Value2 = com.to_excel_serial(period.end)


def write_config(ws, values: dict[str, str]) -> None:
    """Column Value of the Config rows named in column Parameter."""
    rows = {ws.Cells(r, 1).Value: r for r in range(2, ws.UsedRange.Rows.Count + 1)}
    for name, value in values.items():
        ws.Cells(rows[name], 2).Value = com.as_cell(value)


def fill_copy(path: Path, url: str) -> None:
    with com.excel_session() as app:
        wb = com.open_without_macros(app, path)
        write_config(wb.Worksheets("Config"), config_values(url))
        write_periods(wb.Worksheets("Periods"), iterations())
        write_periods(wb.Worksheets("PIs"), PIS)
        wb.Save()
        wb.Close(False)
        del wb, app                                                   # release every COM proxy before Excel quits


EXAMPLES = """
examples:
  scripts\\run.ps1 mock-workbook                                 the copy of out\\release\\jira2dashboard.xlsm
  scripts\\run.ps1 mock-workbook --workbook C:\\x.xlsm --port 8100  another workbook, a fake Jira on port 8100
"""


def parse(argv: list[str] | None = None):
    p = parser("scripts\\run.ps1 mock-workbook", __doc__, EXAMPLES)
    p.add_argument("--workbook", default=str(RELEASE_DIR / "jira2dashboard.xlsm"),
                   help="the workbook to copy (default out\\release\\jira2dashboard.xlsm); it is never changed")
    p.add_argument("--port", type=int, default=8099, help="port of the fake Jira on 127.0.0.1 (default 8099)")
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse(argv)
    source = Path(args.workbook)
    if not source.is_file():
        print(f"mock-workbook: no workbook {source}: run scripts\\run.ps1 build first")
        return 1
    COPY.parent.mkdir(parents=True, exist_ok=True)
    try:
        shutil.copyfile(source, COPY)
    except PermissionError:
        print(f"mock-workbook: close {COPY.name} first (it is still open in Excel)")
        return 1
    fill_copy(COPY, base_url(args.port))
    print(f"mock-workbook: {COPY}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
