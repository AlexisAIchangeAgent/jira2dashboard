"""Build the workbook to deliver: out\\release\\jira2dashboard.xlsm (needs Excel and access to the VBA project).

It first empties out\\inter (scripts\\run.ps1 clean), then builds and compiles the workbook. The workbook holds the VBA that reads Jira and builds the flow sheets and the dashboard (no Python is needed to use
it), the parameter sheets (Config, Projects, Iterations, Workflows, Transitions, FieldMapping) and the plan of the flow
sheets (FlowFormulas, ChartSpec). The client values stay empty: they are typed in the sheet Config after delivery.
Exit code 0 when the VBA compiles, 1 otherwise.
"""
import datetime as dt
import sys
from pathlib import Path

from jira2dashboard import clean, version
from jira2dashboard.cli.args import bare_xlsm, parser
from jira2dashboard.excel import com as R
from jira2dashboard.flow.sheets import write_blueprint, write_manual
from jira2dashboard.paths import INTER_DIR, RELEASE_DIR
from jira2dashboard.params.model import load_all, mandatory_parameters
from jira2dashboard.params.write import write_params
from jira2dashboard.testbench import run_vba_tests as rvt

DELIVERY_NAME = "jira2dashboard.xlsm"


def target_for(name: str) -> Path:
    """out\\release for the workbook to deliver, out\\inter for a variant (--out other.xlsm)."""
    return (RELEASE_DIR if name == DELIVERY_NAME else INTER_DIR) / name


TARGET = target_for(DELIVERY_NAME)
# No test module, no test button: modCoverage stays because the other modules call Cov (a no-op unless switched on).
MODULES = ["modCoverage", "modLog", "modJson", "modText", "modHttp", "modProxy", "modJira", "modFlow", "modParams", "modExtract",
           "modJiraConfig", "modFlowBuild", "modActions"]
BUTTON_MACRO = "modActions.BtnExtractAndCompute"
# The handler of the workbook: a slice of the extraction still scheduled when the workbook is closed would reopen it.
THISWORKBOOK_CODE = "\r\n".join((
    "Option Explicit",
    "Private Sub Workbook_BeforeClose(Cancel As Boolean)",
    "    modActions.AbortExtraction",
    "End Sub",
    ""))


def build_delivery(target: Path = TARGET, release: str | None = None) -> Path:
    """The workbook to deliver at `target`; its hidden name J2D_Version says which commit built it (`release`: the
    version given by scripts\\release.ps1), read back on the RUN line of the logs (version.py)."""
    sources = rvt.module_sources(MODULES)
    stamp = version.stamp(release, *version.commit(), dt.datetime.now())
    with R.excel_session() as app:
        wb = R.new_workbook(app)
        R.assert_vba_access(wb)                                       # fail now, not deep inside the module import
        dash = wb.Worksheets(1)
        dash.Name = "Dashboard"
        params = load_all()
        write_params(wb, params)                                      # Config, Projects, Iterations, Workflows, ...
        write_blueprint(wb, dash, BUTTON_MACRO)
        config = next(sheet for sheet in params if sheet.name == "Config")
        write_manual(wb, dash, mandatory_parameters(config))          # first steps on the Dashboard, sheet Manual
        for name, code in sources.items():
            R.add_or_replace_module(wb, name, code)
        R.set_module_code(R.workbook_component(wb), THISWORKBOOK_CODE)
        wb.Names.Add(version.NAME, version.formula(stamp), False)      # hidden: Name, RefersTo, Visible
        dash.Activate()
        R.save_as_xlsm(wb, target, overwrite=True)
        wb.Close(False)
        del dash, wb, app                                             # release every COM proxy before Excel quits
    return target


EXAMPLES = """
examples:
  scripts\\run.ps1 build                    out\\release\\jira2dashboard.xlsm
  scripts\\run.ps1 build --out try.xlsm     out\\inter\\try.xlsm (a variant, not the one to deliver)
"""


def parse(argv: list[str] | None = None):
    p = parser("scripts\\run.ps1 build", __doc__, EXAMPLES)
    p.add_argument("--out", type=bare_xlsm, default=DELIVERY_NAME, metavar="NAME.xlsm",
                   help=f"name of the workbook: {DELIVERY_NAME} (default) goes to out\\release, any other name to "
                        "out\\inter")
    p.add_argument("--version", metavar="vX.Y.Z", default=None,
                   help="the version of a release, written into the workbook with the commit (scripts\\release.ps1 gives it)")
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse(argv)
    target = target_for(args.out)
    clean.main()                                                       # no intermediates left from older runs
    build_delivery(target, args.version)
    print(f"built {target}")
    result = rvt.compile_ok(target)
    print("compile:", result)
    return 0 if result == "COMPILED_OK" else 1


if __name__ == "__main__":
    sys.exit(main())
