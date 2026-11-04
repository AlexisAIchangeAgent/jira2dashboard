# Python cookbook: every Python file of the kit, ready to extract

Python files cannot be sent by Outlook (`.py` is blocked), and VBA `.bas` files neither. So every file of the
kit is stored in Markdown: this file holds the Python files; `vba-modules-tested.md` holds the VBA modules.
Each file has a heading with its relative path and its MD5. A small bootstrap script recreates them all and
checks every MD5. **Never retype a file by hand or with an AI**: one wrong character breaks it silently.

**Test status:** every file below was run on 2026-11-03 on Windows 11, Python 3.14.7, pywin32 312,
Excel 16.0 Microsoft 365 (French UI): `check_reference.py` 21 checks, 0 failed; `run_vba_tests.py`
143 VBA checks OK (135 offline + 8 network), coverage 65/65 markers = 100 %.

## 1. Where to work

Use a developer folder that the antivirus and IT allow for scripts and macro-enabled files (ask IT).
Do not work in `%TEMP%` or `Downloads`: generated `.xlsm` files and test files may be blocked or deleted there.

## 2. Get all the files (3 steps)

1. Put the Markdown files in `docs/` inside the project folder.
2. Create `tools/cookbook_extract.py` by **copy-paste** of the first block of section 4 (do not retype).
3. From the project folder, run:

```text
py -3.14 tools/cookbook_extract.py docs/cookbook_python.md docs/vba-modules-tested.md
```

Expected: one `OK` line per file (7 Python files, 9 `.bas` files), then `all files OK`.
`MD5 MISMATCH` means the Markdown was changed (for example by an email client or an editor): get a clean
copy of the Markdown file. `EXISTS (use --force)` means a file with another content is already there.

Result (relative to the project folder):

```text
docs/                      the Markdown files
tools/cookbook_extract.py  the bootstrap
xlsm_ref.py                Python helpers (from pywin32-excel-vba-reference.md)
_check_child.py            child: compile + Ping
_hang_child.py             child: timeout test
_run_tests_child.py        child: run the VBA tests
check_reference.py         machine readiness check
run_vba_tests.py           VBA test harness
vba/*.bas                  the 9 tested VBA modules
out/                       created by the scripts (generated workbooks)
```

## 3. Run

Prerequisite: Excel option "Trust access to the VBA project object model" is on
(see `pywin32-excel-vba-reference.md` section 5). In `cmd.exe`, first run `chcp 65001` and
`set PYTHONIOENCODING=utf-8` so French messages print correctly.

```text
py -3.14 check_reference.py                       expected last line: 21 checks, 0 failed
py -3.14 run_vba_tests.py                         offline + network tests
py -3.14 run_vba_tests.py --offline               no network
py -3.14 run_vba_tests.py --out my_tests.xlsm     another output name (use it if the file is open in Excel)
py -3.14 run_vba_tests.py --bisect                find the first module that does not compile
```

Expected output of `run_vba_tests.py` (2026-11-03):

```text
built <project folder>\out\vba_tests.xlsm with 9 modules, 65 coverage markers
compile: COMPILED_OK
summary: 135 OK, 0 KO | net: 8 OK, 0 KO
coverage: 65/65 markers = 100.0 %
missed markers: none
```

The workbook in `out\` has three buttons on `TestResults` (run offline, run with network, clear) and a
progress bar. Without the network (client network, no internet), use `--offline` and the offline button.

## 4. Files

**Bootstrap.** The only file you copy by hand: select the whole block, copy, paste into a new file `tools/cookbook_extract.py`. Do not retype it. Then run it: it writes all the other files and checks every MD5, including its own.

### `tools/cookbook_extract.py`  (MD5 `87ca7a59907135d1d01afaa63dcbe526`)

```python
"""Recreate every file of the kit from the Markdown files, and check each MD5.

Usage, from the project folder (files are written relative to the current folder):
    py -3.14 tools/cookbook_extract.py docs/cookbook_python.md docs/vba-modules-tested.md
    add --force to overwrite files that already exist with another content.
Never retype the files by hand or by an AI: one wrong character breaks them silently.
"""
import hashlib
import re
import sys
from pathlib import Path

FENCE = "`" * 3    # never write three backticks literally here: it would end the Markdown code block
TICK = "`"
SECTION = re.compile(
    r"^### " + TICK + r"([^" + TICK + r"\n]+)" + TICK + r"  \(MD5 " + TICK + r"([0-9a-f]{32})" + TICK + r"\)\n\n"
    + FENCE + r"[a-z]*\n(.*?)\n" + FENCE + r"$", re.S | re.M)


def extract(md_path: Path, force: bool) -> int:
    """Write each file found in one Markdown file; return the number of problems."""
    problems = 0
    for rel, expected, code in SECTION.findall(md_path.read_text(encoding="utf-8")):
        data = (code + "\n").encode("utf-8")
        status = "OK"
        if hashlib.md5(data).hexdigest() != expected:
            status = "MD5 MISMATCH"
        elif rel.endswith(".bas") and not data.isascii():
            status = "NOT ASCII"
        else:
            target = Path(rel)
            if target.exists() and target.read_bytes() != data and not force:
                status = "EXISTS (use --force)"
            else:
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(data)
        problems += status != "OK"
        print(f"{status:22} {rel}")
    return problems


def main(args: list[str]) -> int:
    force = "--force" in args
    files = [Path(a) for a in args if not a.startswith("--")]
    if not files:
        print(__doc__)
        return 2
    problems = sum(extract(f, force) for f in files)
    print("all files OK" if problems == 0 else f"{problems} problem(s)")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
```

All the Python helpers of `pywin32-excel-vba-reference.md`, in one module (the reference's code blocks, in order, without the test fakes and without the child template). Tested: identical to the file that passed `check_reference.py` and `run_vba_tests.py`.

### `xlsm_ref.py`  (MD5 `5707a73a0e99c0835a3791b4f77b64cd`)

```python
"""Code assembled verbatim from pywin32-excel-vba-reference.md (python blocks)."""
import win32api, win32con
import pythoncom                      # COM init, constants for threads
import pywintypes                     # pywintypes.com_error
import win32com.client as win32       # DispatchEx
import win32process                   # GetWindowThreadProcessId (cleanup)
import winreg                         # read the VBA trust setting (standard library)


# Excel
XL_OPEN_XML_WORKBOOK_MACRO_ENABLED = 52   # XlFileFormat: .xlsm
XL_WBAT_WORKSHEET = -4167                 # XlWBATemplate: new workbook with one sheet
XL_SHEET_VISIBLE = -1                     # XlSheetVisibility
XL_SHEET_HIDDEN = 0
XL_SHEET_VERY_HIDDEN = 2                  # hidden, not unhideable from the UI
XL_BUTTON_CONTROL = 0                     # XlFormControl: Form button

# Office AutomationSecurity (MsoAutomationSecurity)
MSO_AUTOMATION_SECURITY_LOW = 1           # macros enabled for files opened by code
MSO_AUTOMATION_SECURITY_BY_UI = 2         # use Trust Center setting
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3 # macros disabled for files opened by code

# VBE component types (vbext_ComponentType)
VBEXT_CT_STD_MODULE = 1
VBEXT_CT_CLASS_MODULE = 2
VBEXT_CT_MS_FORM = 3
VBEXT_CT_DOCUMENT = 100                   # ThisWorkbook and sheet modules (cannot be added)

# HRESULT values (signed 32-bit, as pywin32 reports them)
CO_E_CLASSSTRING = -2147221005            # 0x800401F3: Excel not installed / not registered
RPC_E_CALL_REJECTED = -2147418111         # 0x80010001: Excel busy, retry
RPC_E_SERVERCALL_RETRYLATER = -2147417846 # 0x8001010A: Excel busy, retry
DISP_E_EXCEPTION = -2147352567            # 0x80020009: Excel raised an error, see excepinfo
XL_ERROR_1004 = -2146827284               # 0x800A03EC: Excel generic error 1004 (in excepinfo[5]). Tested.


def com_scode(err: pywintypes.com_error) -> int | None:
    """Return the inner scode of a COM error, or None."""
    info = err.excepinfo
    return info[5] if info else None


import time

BUSY_CODES = {RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER}

def call_with_retry(func, *args, attempts: int = 5, delay_s: float = 0.5):
    """Call a COM function, retrying while Excel is busy."""
    for attempt in range(1, attempts + 1):
        try:
            return func(*args)
        except pywintypes.com_error as err:
            if err.hresult not in BUSY_CODES or attempt == attempts:
                raise
            time.sleep(delay_s)


import contextlib
import gc
import logging

log = logging.getLogger(__name__)

class ExcelNotInstalledError(RuntimeError):
    """Excel desktop is not installed or not registered for COM."""

def start_excel():
    """Start a new hidden Excel process."""
    try:
        app = win32.DispatchEx("Excel.Application")
    except pywintypes.com_error as err:
        if err.hresult == CO_E_CLASSSTRING:
            raise ExcelNotInstalledError("Excel desktop is not installed.") from err
        raise
    app.Visible = False
    app.DisplayAlerts = False
    app.ScreenUpdating = False
    app.EnableEvents = False
    app.AskToUpdateLinks = False
    return app

def excel_pid(app) -> int:
    """Return the process id of this Excel instance."""
    _thread_id, pid = win32process.GetWindowThreadProcessId(app.Hwnd)
    return pid

@contextlib.contextmanager
def excel_session():
    """Yield a hidden Excel app and always close it."""
    pythoncom.CoInitialize()
    try:
        app = start_excel()
        pid = excel_pid(app)
        log.info("Excel %s started, pid %s", app.Version, pid)
        try:
            yield app
        finally:
            _quit_quietly(app)
            del app
            gc.collect()
            _kill_if_still_excel(pid)
    finally:
        pythoncom.CoUninitialize()

def _quit_quietly(app) -> None:
    """Close every workbook without saving, then quit Excel."""
    try:
        while app.Workbooks.Count > 0:
            app.Workbooks(1).Close(False)
        app.Quit()
    except pywintypes.com_error:
        log.warning("Excel did not quit cleanly; it will be killed.")

def _kill_if_still_excel(pid: int, grace_s: float = 5.0) -> None:
    """Wait for Excel to exit; kill it only if it is still EXCEL.EXE after grace_s."""
    access = (win32con.PROCESS_TERMINATE | win32con.PROCESS_QUERY_INFORMATION
              | win32con.PROCESS_VM_READ)
    deadline = time.monotonic() + grace_s
    while True:
        try:
            handle = win32api.OpenProcess(access, False, pid)
        except pywintypes.error:
            return                            # already gone: normal case
        if time.monotonic() >= deadline:
            break
        win32api.CloseHandle(handle)
        time.sleep(0.25)
    try:
        exe = win32process.GetModuleFileNameEx(handle, 0)
        if exe.lower().endswith("excel.exe"):
            win32api.TerminateProcess(handle, 1)
            log.info("Ended Excel pid %s after Quit (references still held)", pid)
    finally:
        win32api.CloseHandle(handle)


USER_KEY = r"Software\Microsoft\Office\16.0\Excel\Security"
POLICY_KEY = r"Software\Policies\Microsoft\Office\16.0\Excel\Security"

def read_access_vbom(key_path: str) -> int | None:
    """Return the AccessVBOM DWORD under HKCU, or None if absent."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            value, _type = winreg.QueryValueEx(key, "AccessVBOM")
            return int(value)
    except FileNotFoundError:
        return None

def is_vbom_trusted() -> bool:
    """True if Excel lets code access VBA projects."""
    policy = read_access_vbom(POLICY_KEY)
    if policy is not None:
        return policy == 1
    return read_access_vbom(USER_KEY) == 1


class VbaAccessNotTrustedError(RuntimeError):
    """Trust access to the VBA project object model is off."""

def assert_vba_access(wb) -> None:
    """Raise a clear error if wb.VBProject cannot be read."""
    try:
        _ = wb.VBProject.VBComponents.Count
    except pywintypes.com_error as err:
        raise VbaAccessNotTrustedError(
            "Turn on: Excel > File > Options > Trust Center > Trust Center Settings > "
            "Macro Settings > Trust access to the VBA project object model."
        ) from err


def new_workbook(app):
    """Create a workbook with exactly one worksheet."""
    return app.Workbooks.Add(XL_WBAT_WORKSHEET)

def add_sheet_after_last(wb, name: str):
    """Add a worksheet at the end and name it."""
    last = wb.Worksheets(wb.Worksheets.Count)
    ws = wb.Worksheets.Add(None, last)        # positional: (Before, After). Tested.
    ws.Name = name
    return ws


def as_cell(value):
    """Pure: give text a leading apostrophe so Excel keeps it as text (tested, see rules below)."""
    if isinstance(value, str) and value:
        return "'" + value[:32766]
    return value

def to_2d_tuple(rows: list[list]) -> tuple[tuple, ...]:
    """Pure: check rows have equal length and return a tuple of tuples."""
    widths = {len(r) for r in rows}
    if len(widths) > 1:
        raise ValueError(f"Rows have different lengths: {sorted(widths)}")
    return tuple(tuple(r) for r in rows)

def block_range(ws, first_row: int, first_col: int, n_rows: int, n_cols: int):
    """Return the range covering a block of cells."""
    return ws.Range(
        ws.Cells(first_row, first_col),
        ws.Cells(first_row + n_rows - 1, first_col + n_cols - 1),
    )

def write_block(ws, first_row: int, first_col: int, rows: list[list]) -> None:
    """Write a 2D list in one COM call."""
    data = to_2d_tuple([[as_cell(v) for v in row] for row in rows])
    if not data:
        return
    block_range(ws, first_row, first_col, len(data), len(data[0])).Value = data


import datetime as dt

_EXCEL_EPOCH = dt.datetime(1899, 12, 30)

def to_excel_serial(value: dt.date | dt.datetime) -> float:
    """Pure: convert a date or naive datetime to an Excel serial number (no time zone shift)."""
    if not isinstance(value, dt.datetime):
        value = dt.datetime(value.year, value.month, value.day)
    if value.tzinfo is not None:
        raise ValueError("Pass a naive (wall-clock) datetime")
    return (value - _EXCEL_EPOCH).total_seconds() / 86400


LCID_EN_US = 0x0409

def put_en_us(obj, prop: str, value) -> None:
    """Set a COM property with the en-US locale (LCID 1033) instead of pywin32's default."""
    ole = obj._oleobj_
    dispid = ole.GetIDsOfNames(prop)
    ole.Invoke(dispid, LCID_EN_US, pythoncom.DISPATCH_PROPERTYPUT, 0, value)


def components(wb):
    """Return the VBComponents collection of a workbook."""
    return wb.VBProject.VBComponents


def get_component(wb, name: str):
    """Return the component called name, or None."""
    comps = components(wb)
    for i in range(1, comps.Count + 1):
        comp = comps.Item(i)
        if comp.Name == name:
            return comp
    return None

def set_module_code(comp, code: str) -> None:
    """Replace all code in a component."""
    module = comp.CodeModule
    if module.CountOfLines > 0:
        module.DeleteLines(1, module.CountOfLines)
    module.AddFromString(code)

def add_or_replace_module(wb, name: str, code: str, comp_type: int = VBEXT_CT_STD_MODULE) -> None:
    """Create the module if needed, then set its code."""
    comp = get_component(wb, name)
    if comp is None:
        comp = components(wb).Add(comp_type)
        comp.Name = name
    set_module_code(comp, code)


def workbook_component(wb):
    """Return the ThisWorkbook component, whatever the language."""
    _ = wb.VBProject.VBComponents.Count       # touch the project so CodeName is filled
    return components(wb).Item(wb.CodeName)

def sheet_component(wb, ws):
    """Return the code module component of a worksheet."""
    _ = wb.VBProject.VBComponents.Count
    code_name = ws.CodeName
    if code_name:
        return components(wb).Item(code_name)
    return _find_document_by_tab_name(wb, ws.Name)

def _find_document_by_tab_name(wb, tab_name: str):
    """Fallback: scan document components for this tab name."""
    comps = components(wb)
    for i in range(1, comps.Count + 1):
        comp = comps.Item(i)
        if comp.Type == VBEXT_CT_DOCUMENT and comp.Properties("Name").Value == tab_name:
            return comp
    raise LookupError(f"No code module for sheet {tab_name!r}")


def set_sheet_code_name(wb, ws, code_name: str) -> None:
    """Give a sheet a language-independent code name, e.g. shData."""
    sheet_component(wb, ws).Name = code_name


import re

_ATTRIBUTE_LINE = re.compile(r"^\s*Attribute\s+VB_", re.IGNORECASE)
_MODULE_NAME = re.compile(r"^[A-Za-z][A-Za-z0-9_]{0,30}$")

def normalize_vba(code: str) -> str:
    """Return VBA text with CRLF endings, no Attribute lines, one final newline."""
    lines = code.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    kept = [ln.rstrip() for ln in lines if not _ATTRIBUTE_LINE.match(ln)]
    while kept and not kept[0]:
        kept.pop(0)
    while kept and not kept[-1]:
        kept.pop()
    return "\r\n".join(kept) + "\r\n"

def is_valid_module_name(name: str) -> bool:
    """True if name is a legal VBA component name."""
    return bool(_MODULE_NAME.match(name))

def vba_string_literal(value: str) -> str:
    """Return a VBA string expression for value (quotes doubled, line breaks as vbLf)."""
    parts = value.replace('"', '""').split("\n")
    return " & vbLf & ".join(f'"{p}"' for p in parts)


MOD_SELF_TEST = r'''
Option Explicit

Public Function Ping() As String
    Ping = "pong"
End Function
'''


def add_macro_button(ws, cell_address: str, caption: str, macro: str,
                     width: float = 120, height: float = 28):
    """Add a Form button at a cell that runs macro (e.g. 'modMain.Run')."""
    cell = ws.Range(cell_address)
    shape = ws.Shapes.AddFormControl(XL_BUTTON_CONTROL, cell.Left, cell.Top, width, height)
    shape.OnAction = macro
    shape.TextFrame.Characters().Text = caption   # Tested
    return shape


from pathlib import Path

def save_as_xlsm(wb, path: Path, overwrite: bool = False) -> Path:
    """Save the workbook as .xlsm at an absolute path."""
    target = path.resolve()
    if target.suffix.lower() != ".xlsm":
        raise ValueError(f"Extension must be .xlsm: {target}")
    if target.exists():
        if not overwrite:
            raise FileExistsError(target)
        target.unlink()
    target.parent.mkdir(parents=True, exist_ok=True)
    wb.SaveAs(str(target), XL_OPEN_XML_WORKBOOK_MACRO_ENABLED)
    return target

def open_without_macros(app, path: Path):
    """Open a workbook with all macros disabled (for checks)."""
    app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
    return app.Workbooks.Open(str(path.resolve()))


COMPILE_CONTROL_ID = 578   # VBE "Compile VBAProject", same ID in every language

def compile_vba_project(app, wb) -> bool:
    """Compile the whole VBA project; True if it compiled without error."""
    app.VBE.ActiveVBProject = wb.VBProject
    ctrl = app.VBE.CommandBars.FindControl(1, COMPILE_CONTROL_ID)
    ctrl.Execute()
    return not ctrl.Enabled

def run_macro(app, wb, macro: str):
    """Run a macro in wb and return its result."""
    return app.Run(f"'{wb.Name}'!{macro}")


import re
import subprocess
import sys

_PID_LINE = re.compile(r"child excel pid (\d+)")

def run_child(script: Path, xlsm_path: Path, *extra_args: str, timeout_s: int = 60) -> str:
    """Run a child check script; on timeout kill its Excel too. Return the last stdout line."""
    try:
        res = subprocess.run(
            [sys.executable, str(script), str(xlsm_path.resolve()), *extra_args],
            capture_output=True, text=True, timeout=timeout_s, check=True,
        )
    except subprocess.TimeoutExpired as err:
        _kill_child_excel(err.stderr)
        raise
    return res.stdout.strip().splitlines()[-1]

def _kill_child_excel(stderr) -> None:
    """Find the child's Excel PID in its stderr and end that process."""
    text = stderr.decode(errors="replace") if isinstance(stderr, bytes) else (stderr or "")
    match = _PID_LINE.search(text)
    if match:
        _kill_if_still_excel(int(match.group(1)), grace_s=0)
```

Child process: opens an `.xlsm`, compiles the whole VBA project, optionally runs `modSelfTest.Ping`, prints `COMPILED_OK`, `COMPILED_OK pong` or `COMPILE_FAILED` (section 12.2 of the pywin32 reference).

### `_check_child.py`  (MD5 `7e2444e3d08c14cd99930240b48c0c80`)

```python
"""Child: open an .xlsm, compile its VBA, optionally run Ping, print the result."""
import sys
from pathlib import Path

from xlsm_ref import (   # adapt to the module that holds the helpers
    MSO_AUTOMATION_SECURITY_FORCE_DISABLE, MSO_AUTOMATION_SECURITY_LOW,
    compile_vba_project, excel_pid, excel_session, run_macro,
)

def main(path: str, run_ping: bool) -> None:
    with excel_session() as app:
        print(f"child excel pid {excel_pid(app)}", file=sys.stderr, flush=True)
        app.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW if run_ping else MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        wb = app.Workbooks.Open(str(Path(path).resolve()))
        ok = compile_vba_project(app, wb)
        result = "COMPILED_OK" if ok else "COMPILE_FAILED"
        if ok and run_ping:
            result += " " + str(run_macro(app, wb, "modSelfTest.Ping"))
        print(result, flush=True)
        wb.Close(False)
        del wb

if __name__ == "__main__":
    main(sys.argv[1], run_ping="--ping" in sys.argv[2:])
```

Child process that hangs on purpose: used by `check_reference.py` to prove that a timeout also ends the child's Excel.

### `_hang_child.py`  (MD5 `bf4f5719d6fb312b1ba1a0d7c35be444`)

```python
"""Child that starts Excel, prints its PID, then hangs: tests the timeout cleanup."""
import sys
import time
from xlsm_ref import excel_pid, excel_session

with excel_session() as app:
    print(f"child excel pid {excel_pid(app)}", file=sys.stderr, flush=True)
    time.sleep(300)
```

**Machine readiness check** (21 checks): Python, pywin32, Excel, VBA trust, sheets, data, dates, number formats, names, modules, code names, events, button, save, reopen, whole-project compile, Ping, compile error detection, timeout cleanup, no Excel left. Run it once on every machine that generates workbooks.

### `check_reference.py`  (MD5 `92ea60e6d9313ba5c5f576d09f637e07`)

```python
"""Run every function of xlsm_ref.py against the real Excel and report each step."""
import datetime as dt
import logging
import subprocess
import sys
import traceback
from pathlib import Path

import xlsm_ref as R

logging.basicConfig(level=logging.INFO, format="  log: %(message)s")
OUT = Path(__file__).with_name("out")
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
    OUT.mkdir(exist_ok=True)
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

    here = Path(__file__).parent
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
```

**VBA test harness**: builds a workbook with the modules of `vba/`, three buttons and a progress bar, compiles it in a child process, runs the VBA tests with coverage, prints KO rows and the coverage.

### `run_vba_tests.py`  (MD5 `4993d05d8ad5ce94bc0fce33120df428`)

```python
"""Build a workbook with the VBA modules in vba/, compile it, run the VBA tests, report results and coverage.

Usage:  py -3.14 run_vba_tests.py            (offline + network tests)
        py -3.14 run_vba_tests.py --offline  (no network tests)
        py -3.14 run_vba_tests.py --bisect   (find the first module that does not compile)
"""
import json
import re
import subprocess
import sys
from pathlib import Path

import xlsm_ref as R

HERE = Path(__file__).parent
VBA_DIR = HERE / "vba"
OUT = HERE / "out" / (sys.argv[sys.argv.index("--out") + 1] if "--out" in sys.argv else "vba_tests.xlsm")
ORDER = ["modCoverage", "modJson", "modText", "modHttp", "modJira", "modFlow", "modBackground", "modTests",
         "modButtons"]
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
        for sheet in ["TestScratch", "CoverageHits", "CoverageMap"]:
            R.add_sheet_after_last(wb, sheet)
        write_coverage_map(wb.Worksheets("CoverageMap"), markers)
        for name, code in sources.items():
            R.add_or_replace_module(wb, name, code)
        if "modButtons" in sources:
            for cell, caption, macro in BUTTONS:
                R.add_macro_button(wb.Worksheets("TestResults"), cell, caption, macro, 190, 26)
            add_progress_bar(wb.Worksheets("TestResults"))
        R.save_as_xlsm(wb, target, overwrite=True)
        wb.Close(False)
        del wb
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
    return R.run_child(HERE / "_check_child.py", target, timeout_s=120)


def bisect() -> None:
    """Add modules one by one and compile each step to find the first failure."""
    for n in range(1, len(ORDER) + 1):
        target = HERE / "out" / f"bisect_{n}.xlsm"
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
    network = "--offline" not in sys.argv
    args = ["--network"] if network else []
    proc = subprocess.run([sys.executable, str(HERE / "_run_tests_child.py"), str(OUT), *args],
                          capture_output=True, text=True, timeout=600)
    if proc.returncode != 0:
        print(proc.stderr[-3000:])
        return 1
    report = json.loads(proc.stdout.strip().splitlines()[-1])
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
```

Child process of `run_vba_tests.py`: opens the test workbook with macros enabled, runs `modTests.RunWithCoverage`, saves, prints a JSON report.

### `_run_tests_child.py`  (MD5 `5a8f81ffdb8d5bfe7d40c4f766bcd672`)

```python
"""Child: open the test workbook with macros enabled, run the VBA tests, save, print a JSON report."""
import json
import sys
from pathlib import Path

import xlsm_ref as R


def main(path: str, network: bool) -> None:
    with R.excel_session() as app:
        print(f"child excel pid {R.excel_pid(app)}", file=sys.stderr, flush=True)
        app.AutomationSecurity = R.MSO_AUTOMATION_SECURITY_LOW
        wb = app.Workbooks.Open(str(Path(path).resolve()))
        summary = app.Run(f"'{wb.Name}'!modTests.RunWithCoverage", network)   # positional argument
        rows = wb.Worksheets("TestResults").UsedRange.Value
        hits = wb.Worksheets("CoverageHits").UsedRange.Value
        wb.Save()
        wb.Close(False)
        del wb
    ko = [list(r[:7]) for r in rows[1:] if r[6] != "OK"]
    hit_ids = [r[0] for r in hits[1:] if r[0]]
    total = len(rows) - 1
    print(json.dumps({"summary": summary, "total": total, "ko": ko, "hits": hit_ids}, default=str))


if __name__ == "__main__":
    main(sys.argv[1], "--network" in sys.argv[2:])
```
