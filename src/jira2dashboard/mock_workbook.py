"""Prepare a copy of a workbook to try it by hand against the fake Jira (scripts\\manual-test.ps1 runs it).

The copy, out\\manual_test\\jira2dashboard4manualtest.xlsm, gets what a client would type: in Config the address,
account, reference date and custom field ids of the fake Jira of `run.ps1 mock-jira`; 14-day iterations in Iterations;
the PI of the known anomalies in PIs. The workbook given is never changed: the one to deliver keeps its client values
empty. out\\manual_test is never emptied by clean, build or vba-tests; the copy is rewritten at each run.
Next to it, rovo_mock_jira.csv: the same configuration as Rovo would answer it (the CSV of the button Load Rovo CSV), to
try that button by hand. The connection (address, account, token, proxy) is not in it: the button never loads it.
"""
import csv
import datetime as dt
import io
import shutil
import sys
from pathlib import Path

from jira2dashboard import rovo_prompt as RP
from jira2dashboard.cli.args import parser
from jira2dashboard.excel import com
from jira2dashboard.flow.model import Iteration
from jira2dashboard.mockjira.anomalies import PI, PI_WINDOW
from jira2dashboard.mockjira.dataset import FieldIds
from jira2dashboard.mockjira.generator import AS_OF, SINCE
from jira2dashboard.mockjira.server import MOCK_EMAIL, MOCK_TOKEN
from jira2dashboard.params.model import load_all
from jira2dashboard.paths import MANUAL_TEST_DIR, RELEASE_DIR
from jira2dashboard.standalone.xlsx_write import Sheet, write_xlsx

COPY_NAME = "jira2dashboard4manualtest.xlsm"
COPY = MANUAL_TEST_DIR / COPY_NAME
CSV_NAME = "rovo_mock_jira.csv"
CSV_NOTE = "configuration of the fake Jira (run.ps1 mock-jira)"
FIELD_IDS = {"StoryPointsField": "story_points", "BusinessValueField": "business_value", "PIField": "pi",
             "TeamField": "team", "SprintField": "sprint", "TargetStartField": "target_start",
             "TargetEndField": "target_end"}          # Config parameter of a data check -> attribute of FieldIds
PIS = [Iteration(PI, *PI_WINDOW)]
ROWS = "A2:C200"                                      # the rows of Iterations and PIs that the copy rewrites


def base_url(port: int) -> str:
    return f"http://127.0.0.1:{port}"


def config_values(url: str, fields: FieldIds = FieldIds(), as_of: dt.datetime = AS_OF) -> dict[str, str]:
    """Parameter -> value of the Config sheet, for the fake Jira at `url`."""
    return {"BaseUrl": url, "Email": MOCK_EMAIL, "ApiToken": MOCK_TOKEN, "AsOf": f"{as_of:%Y-%m-%d %H:%M}",
            "WorkTypeField": fields.work_type} | {name: getattr(fields, attr) for name, attr in FIELD_IDS.items()}


def iterations(first: dt.datetime = SINCE, as_of: dt.datetime = AS_OF, days: int = 14) -> list[Iteration]:
    """I1, I2... of `days` days from `first`, one after the other, the last one holding `as_of`."""
    iterations, start = [], first
    while start < as_of:
        end = start + dt.timedelta(days=days)
        iterations.append(Iteration(f"I{len(iterations) + 1}", start, end))
        start = end
    return iterations


def params_book(url: str, values: dict[str, str] | None = None) -> list[Sheet]:
    """The parameter sheets of params/*.md filled for the fake Jira at `url`, as sheets of an xlsx (no Excel): Config
    with config_values (and `values` over them), the iterations in Iterations and the PI of the anomalies in PIs, as
    dates. What the Python extraction reads from a filled workbook."""
    filled = config_values(url) | (values or {})
    out = []
    for sheet in load_all():
        rows = list(sheet.rows)
        if sheet.name == "FieldMapping":
            continue
        if sheet.name == "Config":
            value = sheet.header.index("Value")
            rows = [row[:value] + (filled.get(row[0], row[value]),) + row[value + 1:] for row in rows]
        elif sheet.name in ("Iterations", "PIs"):
            rows = [(p.name, p.start, p.end) for p in (iterations() if sheet.name == "Iterations" else PIS)]
        out.append(Sheet(sheet.name, sheet.header, rows))
    return out


def write_params_book(path: Path, url: str, values: dict[str, str] | None = None) -> Path:
    return write_xlsx(path, params_book(url, values))


def csv_value(value) -> str:
    """A cell of a parameter sheet as Rovo writes it: a date as yyyy-mm-dd (with the time when it has one)."""
    if isinstance(value, dt.datetime):
        return f"{value:%Y-%m-%d}" if value.time() == dt.time() else f"{value:%Y-%m-%d %H:%M}"
    return "" if value is None else str(value)


def rovo_csv(port: int = 8099) -> str:
    """The answer Rovo would give for the fake Jira, in the CSV the button Load Rovo CSV reads (rovo_prompt.CSV_COLUMNS):
    the Config values of config_values but the connection (rovo_prompt.PROTECTED, never loaded from a CSV), then every row
    of the sheets the prompt asks for (rovo_prompt.SHEET_RULES), with a BOM as Excel writes a UTF-8 CSV."""
    out = io.StringIO()
    writer = csv.writer(out, lineterminator="\r\n")
    writer.writerow(RP.CSV_COLUMNS)
    for name, value in config_values(base_url(port)).items():
        if name not in RP.PROTECTED:
            writer.writerow(("Config", "", name, value, CSV_NOTE))
    for sheet in params_book(base_url(port)):
        if sheet.name in RP.SHEET_RULES:
            for number, row in enumerate(sheet.rows, start=1):
                for column, value in zip(sheet.header, row):
                    writer.writerow((sheet.name, number, column, csv_value(value), ""))
    return "\ufeff" + out.getvalue()


def write_iterations(ws, iterations: list[Iteration]) -> None:
    """Rows Name | Start | End of the sheet Iterations or PIs; the rows of an earlier run are cleared first."""
    ws.Range(ROWS).ClearContents()
    for row, iteration in enumerate(iterations, start=2):
        ws.Cells(row, 1).Value = com.as_cell(iteration.name)
        ws.Cells(row, 2).Value2 = com.to_excel_serial(iteration.start)
        ws.Cells(row, 3).Value2 = com.to_excel_serial(iteration.end)


def write_config(ws, values: dict[str, str]) -> None:
    """Column Value of the Config rows named in column Parameter."""
    rows = {ws.Cells(r, 1).Value: r for r in range(2, ws.UsedRange.Rows.Count + 1)}
    for name, value in values.items():
        ws.Cells(rows[name], 2).Value = com.as_cell(value)


def fill_copy(path: Path, url: str) -> None:
    with com.excel_session() as app:
        wb = com.open_without_macros(app, path)
        write_config(wb.Worksheets("Config"), config_values(url))
        write_iterations(wb.Worksheets("Iterations"), iterations())
        write_iterations(wb.Worksheets("PIs"), PIS)
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
    (COPY.parent / CSV_NAME).write_bytes(rovo_csv(args.port).encode("utf-8"))
    print(f"mock-workbook: {COPY.parent / CSV_NAME} (to try Load Rovo CSV; type the connection by hand)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
