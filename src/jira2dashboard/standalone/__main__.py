"""Extract Jira and compute the flow metrics with Python only, without the VBA: writes jira2dashboard.xlsx.

1. Reads the parameter sheets of the filled workbook (Config, Projects, Iterations, PIs, Workflows, Transitions) as a
   file, without Excel: save the workbook in Excel first. The mandatory values of Config must be filled.
2. Finds the proxy (ProxyServer, or the Windows settings when it is empty), then reads Jira with WinHTTP, as the VBA
   does: the search of the projects (and Filter), page after page, then the changelogs.
3. Computes the data and the results with the Python reference: Issues, StatusIntervals, IssueLinks, ItemFlow,
   DailyLoad, Metrics, PIMetrics, FillRates, DataChecks, and the 9 charts of the workbook as SVG pictures.
4. Writes them as values (no formula, no VBA) in jira2dashboard.xlsx, next to the workbook read (or --out), with the
   sheets Run (what was done) and Errors (the warnings of the run) and the parameters used (ApiToken masked).
Nothing is written when Jira fails or a parameter is missing: the previous file stays as it was.

Log: every step, every failed call to Jira with its status, attempt and what Jira said, every error with its Python
traceback, in the folder logs next to the workbook (logs\\python.log, never the API token nor the Authorization header).
Exit code 0 when the file is written, 1 otherwise.
"""
import datetime as dt
import logging
import tempfile
import logging.handlers
import sys
import time
import zipfile
from pathlib import Path

from jira2dashboard.cli.args import parser
from jira2dashboard.flow import dashboard_spec as D
from jira2dashboard.flow.formulas import CHANGELOG_RAW, INTERVALS_RAW, ISSUES
from jira2dashboard.flow.model import UNCLASSIFIED
from jira2dashboard.params.model import ParamsError
from jira2dashboard.paths import RELEASE_DIR
from jira2dashboard.standalone import jira_api as J
from jira2dashboard.standalone import results as R
from jira2dashboard.standalone.charts_svg import chart_pictures
from jira2dashboard.standalone.http_winhttp import JiraClient, WinHttp, took
from jira2dashboard.standalone.mapping import FieldIds, extracted, interval_rows
from jira2dashboard.standalone.params_book import read_params, run_settings
from jira2dashboard.standalone.proxy import choose_proxy, detect_proxy
from jira2dashboard import sheet_names as S
from jira2dashboard import version
from jira2dashboard.standalone.xlsx_read import NotAWorkbook, defined_name
from jira2dashboard.standalone.xlsx_write import Sheet, XlsxLocked, write_xlsx

LOG_NAME = "python.log"


def default_log(workbook: str) -> Path:
    """The folder logs next to the workbook (it runs from a folder not known in advance), as the VBA log."""
    return Path(workbook).resolve().parent / "logs" / LOG_NAME
OUT_NAME = "jira2dashboard.xlsx"
MASK = "********"
SECRET_PARAMETERS = ("ApiToken",)                  # not copied in clear into the xlsx; the e-mail stays (the existing choice, test_standalone)
log = logging.getLogger("jira2dashboard.extract")
EXAMPLES = """
examples:
  scripts\\run.ps1 extract                                   reads out\\release\\jira2dashboard.xlsm
  scripts\\run.ps1 extract --workbook C:\\Data\\jira2dashboard.xlsm --out C:\\Data\\flow.xlsx
  scripts\\run.ps1 extract --team Alpha --verbose            Flow Distribution of the team Alpha, each call logged
"""


def parse(argv: list[str] | None = None):
    p = parser("scripts\\run.ps1 extract", __doc__, EXAMPLES)
    p.add_argument("--workbook", default=str(RELEASE_DIR / "jira2dashboard.xlsm"),
                   help="the filled workbook whose parameter sheets are read (default out\\release\\jira2dashboard.xlsm)")
    p.add_argument("--out", default="", help=f"the file to write (default: {OUT_NAME} next to the workbook)")
    p.add_argument("--team", default=D.ALL_TEAMS,
                   help=f"the team of the Flow Distribution chart (default {D.ALL_TEAMS!r}, as the Dashboard)")
    p.add_argument("--log", default="", help=f"the log file (default: logs\\{LOG_NAME} next to the workbook)")
    p.add_argument("--verbose", action="store_true", help="also log each call to Jira that worked")
    args = p.parse_args(argv)
    if args.out:                                      # a plain xlsx must never replace the filled workbook (macros, Config)
        out = Path(args.out)
        if out.suffix.lower() != ".xlsx":
            p.error(f"--out must be a file name ending in .xlsx (got {args.out})")
        if out.resolve() == Path(args.workbook).resolve():
            p.error("--out is the workbook itself: the extraction would replace it")
    args.log = args.log or str(default_log(args.workbook))
    return args


class Scrubbed(logging.Formatter):
    """Removes the secrets of the run (API token, Authorization header) from every line, traceback included."""

    def __init__(self, fmt: str):
        super().__init__(fmt)
        self.secrets: list[str] = []

    def format(self, record: logging.LogRecord) -> str:
        text = super().format(record)
        for secret in self.secrets:
            text = text.replace(secret, MASK)
        return text


class Collected(logging.Handler):
    """The warnings and errors of the run, for the sheet Errors."""

    def __init__(self, formatter: Scrubbed):
        super().__init__(logging.WARNING)
        self.rows: list[tuple] = []
        self.scrub = formatter

    def emit(self, record: logging.LogRecord) -> None:
        message = record.getMessage()
        for secret in self.scrub.secrets:
            message = message.replace(secret, MASK)
        moment = dt.datetime.fromtimestamp(record.created, dt.UTC).replace(tzinfo=None)
        self.rows.append((moment, record.levelname, message))


def _handlers(path: Path, verbose: bool):
    formatter = Scrubbed("%(asctime)s %(levelname)s %(message)s")
    path.parent.mkdir(parents=True, exist_ok=True)
    to_file = logging.handlers.RotatingFileHandler(path, maxBytes=1_000_000, backupCount=1, encoding="utf-8")
    to_console = logging.StreamHandler(sys.stdout)
    console = Scrubbed("extract: %(message)s")
    console.secrets = formatter.secrets                 # the same list: a secret added later is masked on both
    to_console.setFormatter(console)
    to_console.setLevel(logging.INFO)
    to_file.setFormatter(formatter)
    collected = Collected(formatter)
    for handler in (to_file, to_console, collected):
        log.addHandler(handler)
    log.setLevel(logging.DEBUG if verbose else logging.INFO)
    log.propagate = False
    return formatter, collected, (to_file, to_console, collected)


def _secrets(token: str, authorization: str) -> list[str]:
    raw = token or ""
    return [s for s in {raw, J.trim_white(raw), authorization, authorization.removeprefix("Basic ")} if len(s) >= 4]


def _masked(sheets) -> list[Sheet]:
    out = []
    for sheet in sheets:
        rows = list(sheet.rows)
        if sheet.name == "Config":
            value = sheet.header.index("Value")
            rows = [r[:value] + (MASK if r[0] in SECRET_PARAMETERS and r[value] else r[value],) + r[value + 1:] for r in rows]
        out.append(Sheet(S.of(sheet.name), sheet.header, rows))
    return out


def parameter_lines(sheets) -> list[str]:
    """The inputs of the run for the log (PARAM lines): each value of Config, the API token as (set) or (empty), then the
    number of rows of each other parameter sheet. The log is complete on purpose (AL's choice), never the token."""
    lines = []
    for sheet in sheets:
        if sheet.name == "Config":
            value = sheet.header.index("Value")
            for row in sheet.rows:
                text = "" if row[value] is None else str(row[value]).strip()
                shown = ("(set)" if row[0] in SECRET_PARAMETERS else text) if text else "(empty)"
                lines.append(f"{row[0]} = {shown}")
        else:
            lines.append(f"sheet {sheet.name} = {len(sheet.rows)} row{'' if len(sheet.rows) == 1 else 's'}")
    return lines


def run(args, formatter: Scrubbed, collected: Collected, sleep=time.sleep, detect=detect_proxy,
        transport=WinHttp) -> Path:
    started = time.monotonic()
    started_at = dt.datetime.now(dt.UTC).replace(tzinfo=None)
    workbook = Path(args.workbook)
    out = Path(args.out) if args.out else workbook.with_name(OUT_NAME)
    log.info("BEGIN parameters of %s", workbook)
    sheets = read_params(workbook)
    for line in parameter_lines(sheets):
        log.info("PARAM %s", line)
    settings = run_settings(sheets)
    c = settings.connection
    authorization = J.basic_auth(c.email, c.token)
    formatter.secrets += _secrets(c.token, authorization)
    base = J.check_base_url(c.base_url)
    jql = J.projects_jql(settings.projects, settings.filter)
    fields = J.search_fields(settings.work_type_field, settings.fields)
    log.info("END parameters: %d iterations, AsOf %s, FlowLoadMode %s%s", len(settings.iterations), settings.as_of,
             settings.mode, took(started))
    step = time.monotonic()
    proxy = choose_proxy(c.proxy_server, c.proxy_bypass, base, detect)
    log.info("END proxy: %s%s%s", proxy.mode, f" ({proxy.server})" if proxy.server else "", took(step))
    client = JiraClient(base, authorization, c.max_retries,
                        transport(c.timeout_s, proxy.server, proxy.bypass, c.windows_auth), sleep=sleep)
    log.info("BEGIN search: %s", jql)
    step = time.monotonic()
    issues = client.search_all(jql, fields)
    log.info("BEGIN changelogs of %d issues", len(issues))
    pages = client.changelogs([i["key"] for i in issues])
    extract = extracted(issues, pages, FieldIds(settings.work_type_field, settings.fields), settings.projects,
                        settings.blocks_link_type)
    log.info("END Jira: %d issues, %d changelog items, %d links%s", len(extract.rows), extract.changelog_items,
             len(extract.links), took(step))
    if extract.unreadable_dates:
        log.warning("J2D-DATA-003 | %d status changes ignored: their date cannot be read | created: %s",
                    extract.unreadable_dates, extract.unreadable_shapes)
    step = time.monotonic()
    intervals = [row for issue in extract.flow for row in interval_rows(issue, settings.as_of, settings.rules)]
    unknown = sorted({row[3] for row in intervals if row[4] == UNCLASSIFIED})
    if unknown:
        log.warning("statuses not in the sheet Workflows (category Unclassified, data check DC08): %s",
                    ", ".join(unknown))
    results = R.compute(extract.flow, settings)
    first, end = R.iteration_span(settings.iterations)
    pictures = chart_pictures(results.metrics, results.teams, [p.name for p in settings.iterations], results.loads,
                              R.days_of(first, end), args.team, results.predictability)
    log.info("END results: %d items, %d metrics rows, %d charts%s", len(results.items), len(results.metrics),
             len(pictures), took(step))
    run_rows = [
        ("Made by", "jira2dashboard, Python path without VBA (scripts\\run.ps1 extract): values, no formula"),
        ("Started (UTC)", started_at), ("Workbook read", str(workbook.resolve())), ("Jira", base), ("JQL", jql),
        ("Proxy", proxy.mode), ("AsOf (UTC)", settings.as_of), ("FlowLoadMode", settings.mode),
        ("Iterations", f"{len(settings.iterations)}: {settings.iterations[0].name} to {settings.iterations[-1].name}"),
        ("Issues", len(extract.rows)), ("Changelog items", extract.changelog_items),
        ("Status intervals", len(intervals)), ("Blocking links", len(extract.links)),
        ("Retries of a call to Jira", client.retries), ("Flow Distribution chart", args.team),
        ("Warnings", f"{len(collected.rows)} (sheet {ERRORS_SHEET})"), ("Log", str(Path(args.log).resolve())),
        ("Duration (s)", round(time.monotonic() - started, 1)),
    ]
    sheets_out = [
        Sheet(RUN_SHEET, ("Item", "Value"), run_rows),
        Sheet(S.of("Charts"), rows=[("The charts of the run, drawn by Python (SVG pictures: Excel 365 shows them).",)],
              pictures=pictures, pictures_per_row=D.CHARTS_PER_ROW),
        Sheet(S.of("Metrics"), R.METRICS, R.metrics_rows(results, settings)),
        Sheet(S.of("PIMetrics"), R.PI_METRICS, R.pimetrics_rows(results, settings)),
        Sheet(S.of("DataChecks"), R.DATA_CHECKS, R.data_check_rows(results.flags, results.fill, settings.checks)),
        Sheet(S.of("FillRates"), R.FILL_RATES, R.fillrates_rows(results.fill)),
        Sheet(S.of("ItemFlow"), R.ITEMFLOW, R.itemflow_rows(results, extract.rows, extract.flow, settings.projects)),
        Sheet(S.of("DailyLoad"), R.DAILY_LOAD, R.dailyload_rows(results, first, end)),
        Sheet(S.of("Issues"), ISSUES, extract.rows),
        Sheet(S.of("StatusIntervals"), INTERVALS_RAW, intervals),
        Sheet(S.of("IssueLinks"), ("Blocker", "Blocked"), extract.links),
        Sheet(S.of("Changelog"), CHANGELOG_RAW, extract.changelog),
        Sheet(ERRORS_SHEET, ("TimeUtc", "Level", "Message"), collected.rows),
    ] + _masked(sheets)
    write_xlsx(out, sheets_out)
    return out


RUN_SHEET = S.of("Run")                     # sheets of this path only: written by the tool, like the calculated ones
ERRORS_SHEET = S.of("Errors")

UNEXPECTED_CODE = "J2D-SYS-099"


def support_code(err: Exception) -> str:
    """The code of an expected error of the run (error_codes.py): the same text as the VBA gives to the same cause."""
    if isinstance(err, (J.JiraError, ParamsError)):
        return err.support_code if isinstance(err, J.JiraError) else err.code
    if isinstance(err, XlsxLocked):
        return "J2D-SYS-010"
    return "J2D-SYS-011"                                # the workbook: not found, not a zip, not a workbook


def main(argv: list[str] | None = None, *, sleep=time.sleep, detect=detect_proxy, transport=WinHttp) -> int:
    args = parse(argv)
    try:
        formatter, collected, handlers = _handlers(Path(args.log), args.verbose)
    except OSError as err:                              # a read-only or locked folder next to the workbook
        args.log = str(Path(tempfile.gettempdir()) / "jira2dashboard" / "logs" / LOG_NAME)
        print(f"extract: cannot write the log there ({err}); the log goes to {args.log}")
        formatter, collected, handlers = _handlers(Path(args.log), args.verbose)
    started = time.monotonic()
    try:
        log.info("RUN extract | Python %s | %s | jira2dashboard %s", sys.version.split()[0], Path(args.workbook).name,
                 defined_name(Path(args.workbook), version.NAME) or version.UNKNOWN)
        out = run(args, formatter, collected, sleep, detect, transport)
        log.info("RUN end ok: %s%s", out, took(started))
        return 0
    except (ParamsError, J.JiraError, XlsxLocked, NotAWorkbook, FileNotFoundError, zipfile.BadZipFile) as err:
        code = support_code(err)
        log.error("%s | %s", code, err)
        log.info("RUN end stopped %s (log: %s)%s", code, args.log, took(started))
        return 1
    except Exception:                                   # noqa: BLE001 - the traceback goes to the log, then exit 1
        log.exception("%s | unexpected error", UNEXPECTED_CODE)
        log.info("RUN end stopped %s (log: %s)%s", UNEXPECTED_CODE, args.log, took(started))
        return 1
    finally:
        for handler in handlers:
            log.removeHandler(handler)
            handler.close()


if __name__ == "__main__":
    sys.exit(main())
