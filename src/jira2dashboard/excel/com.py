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
        try:
            exe = win32process.GetModuleFileNameEx(handle, 0)
        except pywintypes.error:
            log.info("Excel pid %s is already shutting down; left alone.", pid)   # cannot read it: it is going away
            return
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

import os
from jira2dashboard.paths import SRC_DIR

def child_env(base: dict[str, str] | None = None) -> dict[str, str]:
    """Pure: environment for a child process, with src first on PYTHONPATH (no pip install needed)."""
    env = dict(os.environ if base is None else base)
    existing = env.get("PYTHONPATH")
    env["PYTHONPATH"] = str(SRC_DIR) + (os.pathsep + existing if existing else "")
    return env

def run_child(script: Path, xlsm_path: Path, *extra_args: str, timeout_s: int = 60) -> str:
    """Run a child check script; on timeout kill its Excel too. Return the last stdout line."""
    try:
        res = subprocess.run(
            [sys.executable, str(script), str(xlsm_path.resolve()), *extra_args],
            capture_output=True, text=True, timeout=timeout_s, check=True, env=child_env(),
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
