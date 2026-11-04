"""Build a workbook with the VBA modules in vba/, compile it, run the VBA tests, report results and coverage.

Usage:  scripts\\run.ps1 vba-tests            (offline + network tests)
        scripts\\run.ps1 vba-tests --offline  (no network tests)
        scripts\\run.ps1 vba-tests --bisect   (find the first module that does not compile)
        scripts\\run.ps1 vba-tests --out name.xlsm  (workbook name under out\\)
"""
import json
import re
import subprocess
import sys
from pathlib import Path

from jira2dashboard.excel import com as R
from jira2dashboard.params.model import load_all as load_params
from jira2dashboard.params.write import write_params
from jira2dashboard.paths import CHILDREN_DIR, OUT_DIR, VBA_DIR

OUT = OUT_DIR / (sys.argv[sys.argv.index("--out") + 1] if "--out" in sys.argv else "vba_tests.xlsm")
ORDER = ["modCoverage", "modJson", "modText", "modHttp", "modJira", "modFlow", "modParams", "modExtract",
         "modBackground", "modTests", "modButtons"]
BUTTONS = [("J2", "Lancer les tests (hors ligne)", "modButtons.BtnRunAllTests"),
           ("J4", "Lancer les tests + reseau", "modButtons.BtnRunNetworkTests"),
           ("J6", "Effacer les resultats", "modButtons.BtnClearResults")]
MARKER = re.compile(r'Cov "(\w+)"')
MSO_SHAPE_RECTANGLE = 1


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
        for sheet in ["TestScratch", "TestScratch2", "CoverageHits", "CoverageMap"]:
            R.add_sheet_after_last(wb, sheet)
        write_coverage_map(wb.Worksheets("CoverageMap"), markers)
        if "modParams" in sources:
            write_params(wb, load_params())              # sheets Workflows, Transitions, FieldMapping
        for name, code in sources.items():
            R.add_or_replace_module(wb, name, code)
        if "modButtons" in sources:
            for cell, caption, macro in BUTTONS:
                R.add_macro_button(wb.Worksheets("TestResults"), cell, caption, macro, 190, 26)
            add_progress_bar(wb.Worksheets("TestResults"))
        R.save_as_xlsm(wb, target, overwrite=True)
        wb.Close(False)
        del wb, app                      # release every COM proxy before excel_session quits Excel
    return markers


def add_progress_bar(ws, cell_address: str = "J8", width: float = 190, height: float = 14) -> None:
    """Grey background and green fill rectangles, named pbBack / pbFill for modTests.ProgressDraw."""
    cell = ws.Range(cell_address)
    for name, w, rgb in (("pbBack", width, 0xD9D9D9), ("pbFill", 0.1, 0x50B000)):   # RGB is stored as BGR
        shape = ws.Shapes.AddShape(MSO_SHAPE_RECTANGLE, cell.Left, cell.Top, w, height)
        shape.Name = name
        shape.Fill.ForeColor.RGB = rgb
        shape.Line.Visible = 0


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


def run_tests(target: Path, network: bool) -> dict:
    """Open the workbook in a child process, run the VBA tests; return the child's JSON report."""
    args = ["--network"] if network else []
    proc = subprocess.run([sys.executable, str(CHILDREN_DIR / "_run_tests_child.py"), str(target), *args],
                          capture_output=True, text=True, timeout=600, env=R.child_env())
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr[-3000:])
    return json.loads(proc.stdout.strip().splitlines()[-1])


def bisect() -> None:
    """Add modules one by one and compile each step to find the first failure."""
    for n in range(1, len(ORDER) + 1):
        target = OUT_DIR / f"bisect_{n}.xlsm"
        build(ORDER[:n], target)
        result = compile_ok(target)
        print(f"{ORDER[n - 1]:14} -> {result}")
        if result != "COMPILED_OK":
            return


def main() -> int:
    if "--bisect" in sys.argv:
        bisect()
        return 0
    markers = build(ORDER, OUT)
    print(f"built {OUT} with {len(ORDER)} modules, {len(markers)} coverage markers")
    result = compile_ok(OUT)
    print("compile:", result)
    if result != "COMPILED_OK":
        print("Run with --bisect to find the module that does not compile.")
        return 1
    try:
        report = run_tests(OUT, network="--offline" not in sys.argv)
    except RuntimeError as err:
        print(err)
        return 1
    print("summary:", report["summary"])
    for row in report["ko"]:
        print("  KO", " | ".join(str(x) for x in row))
    hit = set(report["hits"])
    missed = [m for m in markers if m not in hit]
    pct = 100 * (len(markers) - len(missed)) / len(markers)
    print(f"coverage: {len(markers) - len(missed)}/{len(markers)} markers = {pct:.1f} %")
    print("missed markers:", ", ".join(missed) or "none")
    return 0 if not report["ko"] else 1


if __name__ == "__main__":
    sys.exit(main())
