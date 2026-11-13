"""Build a workbook with the VBA modules of vba\\, compile it, run the VBA tests, report results and coverage.

It first empties out\\inter (scripts\\run.ps1 clean), then builds out\\inter\\vba_tests.xlsm with every module and the
test buttons, compiles it and runs the tests T_ (offline) and N_ (network) in a separate Excel. Every Cov marker of
the VBA must be hit (100 %). Exit code 0 when everything passes, 1 otherwise. Needs Excel and access to the VBA project.
"""
import json
import subprocess
import re
import sys
import threading
from contextlib import contextmanager
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from jira2dashboard import clean
from jira2dashboard.cli.args import bare_xlsm, parser
from jira2dashboard.excel import com as R
from jira2dashboard.flow.sheets import write_blueprint
from jira2dashboard.mockjira.server import PAC_PATH, PAC_TYPE, pac_script
from jira2dashboard.params.model import load_all as load_params
from jira2dashboard.params.write import write_params
from jira2dashboard.paths import CHILDREN_DIR, INTER_DIR, VBA_DIR

OUT = INTER_DIR / "vba_tests.xlsm"
ORDER = ["modCoverage", "modLog", "modJson", "modText", "modHttp", "modProxy", "modJira", "modFlow", "modParams", "modExtract",
         "modJiraConfig", "modPropose", "modFlowBuild", "modActions", "modBackground", "modTests", "modButtons"]
BUTTONS = [("J2", "Lancer les tests (hors ligne)", "modButtons.BtnRunAllTests"),
           ("J4", "Lancer les tests + reseau", "modButtons.BtnRunNetworkTests"),
           ("J6", "Effacer les resultats", "modButtons.BtnClearResults")]
MARKER = re.compile(r'Cov "(\w+)"')


def module_sources(names: list[str]) -> dict[str, str]:
    """Read each .bas file; refuse non-ASCII source."""
    out = {}
    for name in names:
        text = (VBA_DIR / f"{name}.bas").read_text(encoding="ascii")
        out[name] = R.normalize_vba(text)
    return out


def coverage_markers(sources: dict[str, str]) -> list[str]:
    """Every marker id in the source, in order, without duplicates."""
    seen: list[str] = []
    for text in sources.values():
        for marker in MARKER.findall(text):
            if marker not in seen:
                seen.append(marker)
    return seen


def build(names: list[str], target: Path) -> list[str]:
    """Create the workbook with sheets and modules; return the marker list."""
    sources = module_sources(names)
    markers = coverage_markers(sources)
    with R.excel_session() as app:
        wb = R.new_workbook(app)
        wb.Worksheets(1).Name = "TestResults"
        for sheet in ["TestScratch", "TestScratch2", "TestScratch3", "CoverageHits", "CoverageMap"]:
            R.add_sheet_after_last(wb, sheet)
        write_coverage_map(wb.Worksheets("CoverageMap"), markers)
        if "modParams" in sources:
            write_params(wb, load_params())              # the six parameter sheets
        if "modFlowBuild" in sources:
            write_blueprint(wb)                          # Issues, StatusIntervals, FlowFormulas, ChartSpec, Dashboard
        for name, code in sources.items():
            R.add_or_replace_module(wb, name, code)
        if "modButtons" in sources:
            for cell, caption, macro in BUTTONS:
                R.add_macro_button(wb.Worksheets("TestResults"), cell, caption, macro, 190, 26)
            R.add_progress_bar(wb.Worksheets("TestResults"), "J8")
        R.save_as_xlsm(wb, target, overwrite=True)
        wb.Close(False)
        del wb, app                      # release every COM proxy before excel_session quits Excel
    return markers


def write_coverage_map(ws, markers: list[str]) -> None:
    """MarkerId | Hits (formula) | Covered, plus a total coverage cell in E2."""
    ws.Range("A1:C1").Value = ("MarkerId", "Hits", "Covered")
    for i, marker in enumerate(markers, start=2):
        ws.Cells(i, 1).Value = marker
        ws.Cells(i, 2).Formula = f"=IFERROR(VLOOKUP(A{i},CoverageHits!A:B,2,FALSE),0)"
        ws.Cells(i, 3).Formula = f"=IF(B{i}>0,1,0)"
    last = len(markers) + 1
    ws.Range("E1").Value = "Coverage"
    ws.Range("E2").Formula = f"=IFERROR(SUM(C2:C{last})/COUNTA(A2:A{last}),0)"
    R.put_en_us(ws.Range("E2"), "NumberFormat", "0.0%")   # plain .NumberFormat is read in French
    ws.Columns("A:E").AutoFit()


def compile_ok(target: Path) -> str:
    return R.run_child(CHILDREN_DIR / "_check_child.py", target, timeout_s=120)


PAC_PROXY = "127.0.0.1:9"       # the proxy that the local script gives: T_ProxyScript only reads it, nothing goes there


@contextmanager
def pac_server():
    """A local HTTP server with one proxy script (PAC), for T_ProxyScript: WinHTTP downloads and runs it, as it would
    the script of the Windows settings. Yields its address; it stops when the tests are done."""
    body = pac_script(PAC_PROXY).encode("ascii")

    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):                                               # noqa: N802 - the name http.server calls
            found = self.path == PAC_PATH
            self.send_response(200 if found else 404)
            self.send_header("Content-Type", PAC_TYPE if found else "text/plain")
            self.send_header("Content-Length", str(len(body) if found else 0))
            self.end_headers()
            if found:
                self.wfile.write(body)

        def log_message(self, *args):                                   # quiet: the report is the test results
            pass

    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    try:
        yield f"http://127.0.0.1:{server.server_address[1]}{PAC_PATH}"
    finally:
        server.shutdown()
        server.server_close()


def run_tests(target: Path, network: bool) -> dict:
    """Open the workbook in a child process, run the VBA tests; return the child's JSON report."""
    args = ["--network"] if network else []
    with pac_server() as pac_url:
        last_line = R.run_child(CHILDREN_DIR / "_run_tests_child.py", target, *args, "--pac-url", pac_url,
                                timeout_s=600)                                                  # kills its Excel on timeout
    if not last_line:
        raise RuntimeError("the test child printed no report")
    return json.loads(last_line)


def bisect() -> bool:
    """Add modules one by one and compile each step to find the first failure. True when every step compiles."""
    for n in range(1, len(ORDER) + 1):
        target = INTER_DIR / f"bisect_{n}.xlsm"
        build(ORDER[:n], target)
        result = compile_ok(target)
        print(f"{ORDER[n - 1]:14} -> {result}")
        if result != "COMPILED_OK":
            return False
    return True


EXAMPLES = """
examples:
  scripts\\run.ps1 vba-tests                   offline and network tests
  scripts\\run.ps1 vba-tests --offline         no network test (some markers are then not required)
  scripts\\run.ps1 vba-tests --out try.xlsm    another workbook name, when vba_tests.xlsm is open in Excel
  scripts\\run.ps1 vba-tests --bisect          find the first module that does not compile
"""


def parse(argv: list[str] | None = None):
    p = parser("scripts\\run.ps1 vba-tests", __doc__, EXAMPLES)
    p.add_argument("--offline", action="store_true", help="skip the network tests N_")
    p.add_argument("--bisect", action="store_true",
                   help="add the modules one by one and compile each step, to find the first that does not compile")
    p.add_argument("--out", type=bare_xlsm, default=OUT.name, metavar="NAME.xlsm",
                   help=f"name of the test workbook, in out\\inter (default {OUT.name})")
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse(argv)
    out = INTER_DIR / args.out
    clean.main()                                                       # no intermediates left from older runs
    if args.bisect:
        return 0 if bisect() else 1
    markers = build(ORDER, out)
    print(f"built {out} with {len(ORDER)} modules, {len(markers)} coverage markers")
    result = compile_ok(out)
    print("compile:", result)
    if result != "COMPILED_OK":
        print("Run with --bisect to find the module that does not compile.")
        return 1
    try:
        report = run_tests(out, network=not args.offline)
    except (RuntimeError, subprocess.TimeoutExpired) as err:
        print(err)
        return 1
    print("summary:", report["summary"])
    for row in report["ko"]:
        print("  KO", " | ".join(str(x) for x in row))
    hit = set(report["hits"])
    missed = [m for m in markers if m not in hit]
    pct = 100 * (len(markers) - len(missed)) / max(len(markers), 1)
    print(f"coverage: {len(markers) - len(missed)}/{len(markers)} markers = {pct:.1f} %")
    print("missed markers:", ", ".join(missed) or "none")
    if report["total"] == 0 or not markers:
        print("no test result or no marker: nothing was tested")
        return 1
    if missed and not args.offline:                  # offline runs skip the network tests, which cover some markers
        print("every marker must be hit (100 % required)")
        return 1
    return 0 if not report["ko"] else 1


if __name__ == "__main__":
    sys.exit(main())
