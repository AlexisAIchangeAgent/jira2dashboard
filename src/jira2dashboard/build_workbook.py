"""Build the workbook to deliver: out\\jira2dashboard.xlsm.

Usage:  scripts\\run.ps1 build [--out name.xlsm]

The workbook holds the VBA that reads Jira and builds the flow sheets and the dashboard (no Python is needed to use
it), the parameter sheets (Config, Projects, Periods, Workflows, Transitions, FieldMapping) and the plan of the flow
sheets (FlowFormulas, ChartSpec). The client values stay empty: they are typed in the sheet Config after delivery.
"""
import sys
from pathlib import Path

from jira2dashboard.excel import com as R
from jira2dashboard.flow.sheets import write_blueprint
from jira2dashboard.paths import OUT_DIR
from jira2dashboard.params.model import load_all
from jira2dashboard.params.write import write_params
from jira2dashboard.testbench import run_vba_tests as rvt

TARGET = OUT_DIR / rvt.option("--out", "jira2dashboard.xlsm")
# No test module, no test button: modCoverage stays because the other modules call Cov (a no-op unless switched on).
MODULES = ["modCoverage", "modJson", "modText", "modHttp", "modJira", "modFlow", "modParams", "modExtract",
           "modFlowBuild", "modActions"]
BUTTON_MACRO = "modActions.BtnExtractAndCompute"


def build_delivery(target: Path = TARGET) -> Path:
    sources = rvt.module_sources(MODULES)
    with R.excel_session() as app:
        wb = R.new_workbook(app)
        dash = wb.Worksheets(1)
        dash.Name = "Dashboard"
        write_params(wb, load_all())                                  # Config, Projects, Periods, Workflows, ...
        write_blueprint(wb, dash, BUTTON_MACRO)
        for name, code in sources.items():
            R.add_or_replace_module(wb, name, code)
        dash.Activate()
        R.save_as_xlsm(wb, target, overwrite=True)
        wb.Close(False)
        del dash, wb, app                                             # release every COM proxy before Excel quits
    return target


def main() -> int:
    target = build_delivery()
    print(f"built {target}")
    result = rvt.compile_ok(target)
    print("compile:", result)
    return 0 if result == "COMPILED_OK" else 1


if __name__ == "__main__":
    sys.exit(main())
