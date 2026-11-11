"""Extract Jira and compute the flow metrics with Python only, without the VBA: writes jira2dashboard.xlsx.

1. Reads the parameter sheets of the filled workbook (Config, Projects, Iterations, PIs, Workflows, Transitions) as a
   file, without Excel: save the workbook in Excel first. The mandatory values of Config must be filled.
2. Finds the proxy (ProxyServer, or the Windows settings when it is empty), then reads Jira with WinHTTP, as the VBA
   does: the search of the projects (and Filter), page after page, then the changelogs.
3. Computes the data and the results with the Python reference: Issues, StatusIntervals, IssueLinks, ItemFlow,
   DailyLoad, Metrics, FillRates, DataChecks, and the 8 charts of the Dashboard as SVG pictures.
4. Writes them as values (no formula, no VBA) in jira2dashboard.xlsx, next to the workbook read (or --out), with the
   sheets Run (what was done) and Errors (the warnings of the run) and the parameters used (ApiToken masked).
Nothing is written when Jira fails or a parameter is missing: the previous file stays as it was.

Log: every step, every failed call to Jira with its status, attempt and what Jira said, every error with its Python
traceback, in the folder logs next to the workbook (logs\\python.log, never the API token nor the Authorization header).
Exit code 0 when the file is written, 1 otherwise.
"""
import datetime as dt
import logging
import logging.handlers
import sys
import time
import zipfile
from pathlib import Path

from jira2dashboard.cli.args import parser
from jira2dashboard.flow import dashboard_spec as D
from jira2dashboard.flow.formulas import INTERVALS_RAW, ISSUES
from jira2dashboard.flow.model import UNCLASSIFIED
from jira2dashboard.params.model import ParamsError
from jira2dashboard.paths import RELEASE_DIR
from jira2dashboard.standalone import jira_api as J
from jira2dashboard.standalone import results as R
from jira2dashboard.standalone.charts_svg import chart_pictures
from jira2dashboard.standalone.http_winhttp import JiraClient, WinHttp
from jira2dashboard.standalone.mapping import FieldIds, extracted, interval_rows
from jira2dashboard.standalone.params_book import read_params, run_settings
from jira2dashboard.standalone.proxy import choose_proxy, detect_proxy
from jira2dashboard.standalone.xlsx_read import NotAWorkbook
from jira2dashboard.standalone.xlsx_write import Sheet, XlsxLocked, write_xlsx

LOG_NAME = "python.log"


def default_log(workbook: str) -> Path:
    """The folder logs next to the workbook (it runs from a folder not known in advance), as the VBA log."""
    return Path(workbook).resolve().parent / "logs" / LOG_NAME
OUT_NAME = "jira2dashboard.xlsx"
MASK = "********"
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
            rows = [r[:value] + (MASK if r[0] == "ApiToken" and r[value] else r[value],) + r[value + 1:] for r in rows]
        out.append(Sheet(sheet.name, sheet.header, rows))
    return out


def run(args, formatter: Scrubbed, collected: Collected, sleep=time.sleep, detect=detect_proxy,
        transport=WinHttp) -> Path:
    started = time.monotonic()
    started_at = dt.datetime.now(dt.UTC).replace(tzinfo=None)
    workbook = Path(args.workbook)
    out = Path(args.out) if args.out else workbook.with_name(OUT_NAME)
    log.info("BEGIN parameters of %s", workbook)
    sheets = read_params(workbook)
    settings = run_settings(sheets)
    c = settings.connection
    authorization = J.basic_auth(c.email, c.token)
    formatter.secrets += _secrets(c.token, authorization)
    base = J.check_base_url(c.base_url)
    jql = J.projects_jql(settings.projects, settings.filter)
    fields = J.search_fields(settings.work_type_field, settings.fields)
    log.info("END parameters: %d iterations, AsOf %s, FlowLoadMode %s", len(settings.iterations), settings.as_of,
             settings.mode)
    proxy = choose_proxy(c.proxy_server, c.proxy_bypass, base, detect)
    log.info("END proxy: %s%s", proxy.mode, f" ({proxy.server})" if proxy.server else "")
    client = JiraClient(base, authorization, c.max_retries,
                        transport(c.timeout_s, proxy.server, proxy.bypass, c.windows_auth), sleep=sleep)
    log.info("BEGIN search: %s", jql)
    issues = client.search_all(jql, fields)
    log.info("BEGIN changelogs of %d issues", len(issues))
    pages = client.changelogs([i["key"] for i in issues])
    extract = extracted(issues, pages, FieldIds(settings.work_type_field, settings.fields), settings.projects,
                        settings.blocks_link_type)
    log.info("END Jira: %d issues, %d changelog items, %d links", len(extract.rows), extract.changelog_items,
             len(extract.links))
    intervals = [row for issue in extract.flow for row in interval_rows(issue, settings.as_of, settings.rules)]
    unknown = sorted({row[3] for row in intervals if row[4] == UNCLASSIFIED})
    if unknown:
        log.warning("statuses not in the sheet Workflows (category Unclassified, data check DC08): %s",
                    ", ".join(unknown))
    results = R.compute(extract.flow, settings)
    first, end = settings.iterations[0].start, settings.iterations[-1].end
    pictures = chart_pictures(results.metrics, results.teams, [p.name for p in settings.iterations], results.loads,
                              R.days_of(first, end), args.team)
    log.info("END results: %d items, %d metrics rows, %d charts", len(results.items), len(results.metrics),
             len(pictures))
    run_rows = [
        ("Made by", "jira2dashboard, Python path without VBA (scripts\\run.ps1 extract): values, no formula"),
        ("Started (UTC)", started_at), ("Workbook read", str(workbook.resolve())), ("Jira", base), ("JQL", jql),
        ("Proxy", proxy.mode), ("AsOf (UTC)", settings.as_of), ("FlowLoadMode", settings.mode),
        ("Iterations", f"{len(settings.iterations)}: {settings.iterations[0].name} to {settings.iterations[-1].name}"),
        ("Issues", len(extract.rows)), ("Changelog items", extract.changelog_items),
        ("Status intervals", len(intervals)), ("Blocking links", len(extract.links)),
        ("Retries of a call to Jira", client.retries), ("Flow Distribution chart", args.team),
        ("Warnings", f"{len(collected.rows)} (sheet Errors)"), ("Log", str(Path(args.log).resolve())),
        ("Duration (s)", round(time.monotonic() - started, 1)),
    ]
    sheets_out = [
        Sheet("Run", ("Item", "Value"), run_rows),
        Sheet("Charts", rows=[("The charts of the run, drawn by Python (SVG pictures: Excel 365 shows them).",)],
              pictures=pictures, pictures_per_row=D.CHARTS_PER_ROW),
        Sheet("Metrics", R.METRICS, R.metrics_rows(results, settings)),
        Sheet("DataChecks", R.DATA_CHECKS, R.data_check_rows(results.flags, results.fill, settings.checks)),
        Sheet("FillRates", R.FILL_RATES, R.fillrates_rows(results.fill)),
        Sheet("ItemFlow", R.ITEMFLOW, R.itemflow_rows(results, extract.rows, extract.flow, settings.projects)),
        Sheet("DailyLoad", R.DAILY_LOAD, R.dailyload_rows(results, first, end)),
        Sheet("Issues", ISSUES, extract.rows),
        Sheet("StatusIntervals", INTERVALS_RAW, intervals),
        Sheet("IssueLinks", ("Blocker", "Blocked"), extract.links),
        Sheet("Errors", ("TimeUtc", "Level", "Message"), collected.rows),
    ] + _masked(sheets)
    write_xlsx(out, sheets_out)
    return out


def main(argv: list[str] | None = None, *, sleep=time.sleep, detect=detect_proxy, transport=WinHttp) -> int:
    args = parse(argv)
    formatter, collected, handlers = _handlers(Path(args.log), args.verbose)
    try:
        log.info("RUN extract | Python %s | %s", sys.version.split()[0], Path(args.workbook).name)
        out = run(args, formatter, collected, sleep, detect, transport)
        log.info("RUN end ok: %s", out)
        return 0
    except (ParamsError, J.JiraError, XlsxLocked, NotAWorkbook, FileNotFoundError, zipfile.BadZipFile) as err:
        log.error("%s", err)
        log.info("RUN end stopped (log: %s)", args.log)
        return 1
    except Exception:                                   # noqa: BLE001 - the traceback goes to the log, then exit 1
        log.exception("unexpected error")
        log.info("RUN end stopped (log: %s)", args.log)
        return 1
    finally:
        for handler in handlers:
            log.removeHandler(handler)
            handler.close()


if __name__ == "__main__":
    sys.exit(main())
