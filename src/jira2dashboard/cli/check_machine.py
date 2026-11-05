"""Run every function of excel/com.py against the real Excel and report each step."""
import datetime as dt
import logging
import subprocess
import sys
import traceback
from pathlib import Path

from jira2dashboard.excel import com as R
from jira2dashboard.paths import CHILDREN_DIR, INTER_DIR

logging.basicConfig(level=logging.INFO, format="  log: %(message)s")
OUT = INTER_DIR
RESULTS: list[tuple[str, str, str]] = []

MOD_MAIN = r'''
Attribute VB_Name = "modMain"
Option Explicit

Public Sub Run()
    shData.Range("D1").Value = "run ok " & ThisWorkbook.Names("cfg_Region").RefersToRange.Value
End Sub

Public Function Greeting() As String
    Greeting = "Caf" & ChrW(233)
End Function
'''

WB_OPEN = r'''
Option Explicit

Private Sub Workbook_Open()
    shData.Range("E1").Value = "opened"
End Sub
'''


def step(name):
    """Decorator: run a check, record PASS/FAIL, never stop the run."""
    def wrap(func):
        def inner(*a, **k):
            try:
                detail = func(*a, **k)
                RESULTS.append(("PASS", name, str(detail or "")))
                print(f"PASS  {name}: {detail or ''}")
                return detail
            except Exception as err:  # noqa: BLE001 - report everything
                info = describe(err)
                RESULTS.append(("FAIL", name, info))
                print(f"FAIL  {name}: {info}")
                traceback.print_exc(limit=2)
                return None
        return inner
    return wrap


def describe(err: Exception) -> str:
    """Short text for an exception, with COM codes when present."""
    if isinstance(err, R.pywintypes.com_error):
        return (f"com_error hresult={err.hresult} scode={R.com_scode(err)} "
                f"text={err.excepinfo[2] if err.excepinfo else err.strerror!r}")
    return f"{type(err).__name__}: {err}"


def skip(name, why):
    RESULTS.append(("SKIP", name, why))
    print(f"SKIP  {name}: {why}")


def excel_count() -> int:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"],
                         capture_output=True, text=True).stdout
    return out.lower().count("excel.exe")


@step("pure: normalize_vba strips Attribute, uses CRLF")
def check_pure():
    code = R.normalize_vba(MOD_MAIN)
    assert "Attribute" not in code and "\r\n" in code and "\n\n" not in code.replace("\r\n", "")
    assert R.vba_string_literal('a"b\nc') == '"a""b" & vbLf & "c"'
    assert R.is_valid_module_name("modMain") and not R.is_valid_module_name("1x")
    return "3 pure helpers ok"


@step("registry: is_vbom_trusted")
def check_registry():
    return f"trusted={R.is_vbom_trusted()} user={R.read_access_vbom(R.USER_KEY)} policy={R.read_access_vbom(R.POLICY_KEY)}"


def build_and_check(app, trusted: bool) -> Path | None:
    ctx = {}

    @step("excel: start hidden instance")
    def s_start():
        return f"Version={app.Version} Visible={app.Visible} EnableEvents={app.EnableEvents} pid={R.excel_pid(app)}"
    s_start()

    @step("workbook: Workbooks.Add(-4167) gives one sheet")
    def s_new():
        wb = R.new_workbook(app)
        ctx["wb"] = wb
        first = wb.Worksheets(1)
        n, name = wb.Worksheets.Count, first.Name
        first.Name = "Data"
        assert n == 1
        return f"count={n}, default name={name!r} renamed to 'Data'"
    s_new()
    wb = ctx.get("wb")
    if wb is None:
        return None
    ws = wb.Worksheets("Data")

    @step("workbook: add_sheet_after_last + very hidden")
    def s_sheet():
        cfg = R.add_sheet_after_last(wb, "Config")
        cfg.Range("A2").Value = "Region"
        cfg.Range("B2").Value = "Nord"
        cfg.Visible = R.XL_SHEET_VERY_HIDDEN
        ctx["cfg"] = cfg
        order = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]
        assert order == ["Data", "Config"], order
        return f"order={[wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]} visible={cfg.Visible}"
    s_sheet()

    @step("data: write_block numbers, None, datetime")
    def s_data():
        day = dt.date(2026, 11, 3)
        rows = [["Item", "Qty", "Date"], ["A", 2, R.to_excel_serial(day)], ["B", 3.5, None]]
        R.write_block(ws, 1, 1, rows)
        R.put_en_us(ws.Range("C2"), "NumberFormat", "yyyy-mm-dd")
        back = R.block_range(ws, 1, 1, 3, 3).Value
        assert back[1][1] == 2 and back[2][1] == 3.5 and back[2][2] is None
        assert back[1][2].replace(tzinfo=None) == dt.datetime(2026, 11, 3), back[1][2]
        aware = dt.datetime(2026, 11, 3, tzinfo=dt.timezone.utc)
        ws.Range("C3").Value = aware
        assert ws.Range("C3").Value2 == ws.Range("C2").Value2
        return f"serial={ws.Range('C2').Value2} read={back[1][2]!r}; UTC-aware gives same serial"
    s_data()

    @step("data: English Formula / NumberFormat in this Excel language")
    def s_formula():
        c = ws.Range("B4")
        c.Formula = "=SUM(B2:B3)"
        R.put_en_us(c, "NumberFormat", "0.00")
        assert c.Value == 5.5
        assert c.NumberFormatLocal in ("0,00", "0.00"), c.NumberFormatLocal   # French "0,00", English "0.00"
        return f"Value={c.Value} Formula={c.Formula!r} FormulaLocal={c.FormulaLocal!r} NumberFormatLocal={c.NumberFormatLocal!r} Text={c.Text!r}"
    s_formula()

    @step("names: Names.Add cfg_Region")
    def s_names():
        wb.Names.Add("cfg_Region", "=Config!$B$2")
        return f"RefersTo={wb.Names('cfg_Region').RefersTo!r} value={wb.Names('cfg_Region').RefersToRange.Value!r}"
    s_names()

    @step("vba: assert_vba_access")
    def s_access():
        R.assert_vba_access(wb)
        return "VBProject readable"
    if not trusted:
        try:
            R.assert_vba_access(wb)
            RESULTS.append(("FAIL", "vba: access refused as expected", "access worked although registry says not trusted"))
        except R.VbaAccessNotTrustedError as err:
            cause = err.__cause__
            RESULTS.append(("PASS", "vba: clear error when not trusted", describe(cause)))
            print(f"PASS  vba: clear error when not trusted: {describe(cause)}")
        for n in ["vba: CodeName quirk", "vba: modules", "vba: code names", "vba: ThisWorkbook event", "button"]:
            skip(n, "VBA trust is off")
    else:
        s_access()
        vba_steps(wb, ws, ctx)

    target = OUT / "check.xlsm"

    @step("save: save_as_xlsm, FileFormat 52")
    def s_save():
        path = R.save_as_xlsm(wb, target, overwrite=True)
        assert wb.FileFormat == R.XL_OPEN_XML_WORKBOOK_MACRO_ENABLED
        return f"{path} FileFormat={wb.FileFormat}"
    s_save()

    @step("save: refuses overwrite by default")
    def s_noover():
        try:
            R.save_as_xlsm(wb, target)
        except FileExistsError:
            return "FileExistsError raised"
        raise AssertionError("no error")
    s_noover()
    wb.Close(False)
    return target


def vba_steps(wb, ws, ctx):
    @step("vba: CodeName quirk on a new sheet")
    def s_quirk():
        tmp = R.add_sheet_after_last(wb, "Tmp")
        before = tmp.CodeName
        _ = wb.VBProject.VBComponents.Count
        after = tmp.CodeName
        comp = R.sheet_component(wb, tmp)
        name = comp.Name
        tmp.Delete()
        return f"CodeName before touch={before!r} after={after!r}; sheet_component -> {name!r}"
    s_quirk()

    @step("vba: add_or_replace_module (twice, Option Explicit dedupe)")
    def s_mod():
        R.add_or_replace_module(wb, "modMain", R.normalize_vba("Sub Old()\r\nEnd Sub"))
        R.add_or_replace_module(wb, "modMain", R.normalize_vba(MOD_MAIN))
        R.add_or_replace_module(wb, "modSelfTest", R.normalize_vba(R.MOD_SELF_TEST))
        cm = R.get_component(wb, "modMain").CodeModule
        text = cm.Lines(1, cm.CountOfLines)
        assert text.lower().count("option explicit") == 1, text
        assert "Old" not in text
        return f"modMain lines={cm.CountOfLines}, Option Explicit count=1"
    s_mod()

    @step("vba: code names shData / shConfig + ThisWorkbook lookup")
    def s_codenames():
        R.set_sheet_code_name(wb, ws, "shData")
        R.set_sheet_code_name(wb, ctx["cfg"], "shConfig")
        wbc = R.workbook_component(wb)
        return f"Data.CodeName={ws.CodeName} Config.CodeName={ctx['cfg'].CodeName} ThisWorkbook comp={wbc.Name!r}"
    s_codenames()

    @step("vba: ThisWorkbook Workbook_Open code")
    def s_event():
        R.set_module_code(R.workbook_component(wb), R.normalize_vba(WB_OPEN))
        cm = R.workbook_component(wb).CodeModule
        return f"lines={cm.CountOfLines}"
    s_event()

    @step("button: add_macro_button with caption")
    def s_button():
        shp = R.add_macro_button(ws, "B6", "Lancer", "modMain.Run")
        return f"OnAction={shp.OnAction!r} caption={shp.TextFrame.Characters().Text!r}"
    s_button()


def reopen_checks(app, path: Path):
    @step("reopen: open_without_macros, read code back")
    def s_reopen():
        wb = R.open_without_macros(app, path)
        try:
            src = R.normalize_vba(MOD_MAIN)
            cm = R.get_component(wb, "modMain").CodeModule
            back = R.normalize_vba(cm.Lines(1, cm.CountOfLines))
            same = back == src
            same_ci = back.lower() == src.lower()
            btn = wb.Worksheets("Data").Shapes(1).OnAction
            return f"FileFormat={wb.FileFormat} exact={same} case-insensitive={same_ci} button OnAction={btn!r} AutomationSecurity={app.AutomationSecurity}"
        finally:
            wb.Close(False)
    s_reopen()


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    base = excel_count()
    print(f"Python {sys.version.split()[0]}, Excel processes before: {base}")
    check_pure()
    check_registry()
    trusted = R.is_vbom_trusted()
    pid = None
    path = None
    try:
        with R.excel_session() as app:
            pid = R.excel_pid(app)
            path = build_and_check(app, trusted)
            if path and trusted:
                reopen_checks(app, path)
    except Exception as err:  # noqa: BLE001
        RESULTS.append(("FAIL", "excel_session", describe(err)))
        traceback.print_exc()

    here = CHILDREN_DIR
    if path and trusted:
        @step("child: compile whole project + Ping (good file)")
        def s_ping():
            out = R.run_child(here / "_check_child.py", path, "--ping", timeout_s=90)
            assert out == "COMPILED_OK pong", out
            return out
        s_ping()

        @step("child: compile detects error in another module")
        def s_bad():
            bad = OUT / "compile_error.xlsm"
            with R.excel_session() as app2:
                wb2 = R.new_workbook(app2)
                R.add_or_replace_module(wb2, "modBad", "Option Explicit\r\nSub Broken()\r\n    x = 1\r\nEnd Sub\r\n")
                R.add_or_replace_module(wb2, "modSelfTest", R.normalize_vba(R.MOD_SELF_TEST))
                R.save_as_xlsm(wb2, bad, overwrite=True)
                wb2.Close(False)
                del wb2
            out = R.run_child(here / "_check_child.py", bad, "--ping", timeout_s=90)
            assert out == "COMPILE_FAILED", out
            return out
        s_bad()

        @step("child: timeout kills the child's Excel")
        def s_timeout():
            before = excel_count()
            try:
                R.run_child(here / "_hang_child.py", path, timeout_s=15)
            except subprocess.TimeoutExpired:
                import time; time.sleep(1)
                after = excel_count()
                assert after <= before, f"before={before} after={after}"
                return f"TimeoutExpired raised, Excel count before={before} after={after}"
            raise AssertionError("no timeout")
        s_timeout()
    else:
        skip("child checks", "VBA trust is off")

    @step("cleanup: no Excel left from this run")
    def s_clean():
        now = excel_count()
        assert now == base, f"before={base} after={now}"
        return f"before={base} after={now} (session pid {pid})"
    s_clean()

    print("\n=== SUMMARY ===")
    for status, name, detail in RESULTS:
        print(f"{status:4}  {name}")
    fails = sum(1 for s, *_ in RESULTS if s == "FAIL")
    print(f"{len(RESULTS)} checks, {fails} failed")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
