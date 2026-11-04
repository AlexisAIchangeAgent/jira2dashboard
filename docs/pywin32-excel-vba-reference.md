# Offline reference: pywin32 312 + Excel COM + VBA projects

This file is the API source for generating `.xlsm` files with VBA from Python.
Target: Windows 11, Python 3.14 x64, pywin32 312, Excel Microsoft 365 desktop (French or English).

**Test status:** every Python block in this file was run on 2026-11-03 on Windows 11,
Python 3.14.7 (64-bit), pywin32 312, Excel 16.0 Microsoft 365, **French** UI.
Result: all checks pass (sheets, data, dates, formulas, names, VBA trust check, modules,
code names, ThisWorkbook event, button, save, reopen, whole-project compile, `Ping` in a child,
Excel cleanup). Not yet run on an **English** Excel: the rules below are written to be
language-independent, but run the check once there too.

Labels used below:
- **Tested**: run on the machine above, result shown.
- **Documented**: matches Microsoft or pywin32 documentation.
- **Known pattern**: widely used, not in one official page.
- **Verify on first run**: likely right, not yet run. Check it once on the target machine.

---

## 1. Install and import

```text
py -3.14 -m pip install pywin32==312 pytest
```

pywin32 312 ships a `cp314-win_amd64` wheel. No post-install step is needed for COM client use.

```python
import pythoncom                      # COM init, constants for threads
import pywintypes                     # pywintypes.com_error
import win32com.client as win32       # DispatchEx
import win32process                   # GetWindowThreadProcessId (cleanup)
import winreg                         # read the VBA trust setting (standard library)
```

**Rule:** use late binding with `win32.DispatchEx("Excel.Application")`.
- `DispatchEx` starts a **new** Excel process. `Dispatch` may attach to the user's open Excel.
- Avoid `win32.gencache.EnsureDispatch`. It writes a cache in `%TEMP%\gen_py`.
  A broken cache gives `AttributeError` errors that are hard to understand.
- With late binding, `win32com.client.constants` is empty. Use the named values in section 2.

Python 64-bit can drive Excel 32-bit or 64-bit. Excel runs in its own process.

---

## 2. Constants (Documented)

```python
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
```

---

## 3. Errors: `pywintypes.com_error`

Every failed COM call raises `pywintypes.com_error`. Its fields:

| Field | Meaning |
|---|---|
| `e.hresult` | Main HRESULT (int, signed) |
| `e.strerror` | Short text (localized, do not test it) |
| `e.excepinfo` | `None` or a tuple: `(wcode, source, description, helpfile, helpcontext, scode)` |
| `e.argerror` | Index of the bad argument, or `None` |

```python
def com_scode(err: pywintypes.com_error) -> int | None:
    """Return the inner scode of a COM error, or None."""
    info = err.excepinfo
    return info[5] if info else None
```

**Console note (tested):** in `cmd.exe`, French messages print with broken accents.
Run with `set PYTHONIOENCODING=utf-8` and `chcp 65001`, or write logs to a UTF-8 file.

**Rule:** test `e.hresult` and `com_scode(e)`. Never test the message text. French Excel gives French text.

Retry helper for "Excel busy" errors (Known pattern):

```python
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
```

---

## 4. Start and stop Excel safely

```python
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
```

Extra imports for this section: `import win32api, win32con` (both in pywin32).

Notes:
- **Tested (Excel 16.0, M365, French):** after `Quit()`, Excel exits about 4 seconds after the last
  COM reference is released. While the caller's `with excel_session() as app:` variable exists,
  Excel cannot exit at all, so the fallback ends it after `grace_s` (5 s). This is safe: `Quit()` was
  sent and every workbook was closed without saving. The log line is normal, not an error.
- Excel stays alive while any Python variable still holds a COM object (`wb`, `ws`, a range).
  Keep COM objects local to functions so they are released when the function returns.
- Only kill the PID you started, and only if it is still `EXCEL.EXE` (a PID can be reused).
  Never kill all `EXCEL.EXE` processes: the user may have work open.
- `CoInitialize` is needed in each thread that uses COM. In the main thread it is harmless.
- `app.Hwnd` is Documented on `Application`. It is the main window handle.

---

## 5. Check "Trust access to the VBA project object model"

Without this setting, any access to `wb.VBProject` raises `com_error`
(`hresult == DISP_E_EXCEPTION`, `scode == XL_ERROR_1004`). **Tested:** French Excel gives the text
"L'accès par programme au projet Visual Basic n'est pas fiable". Do not match on it.
The setting is per Windows user, on the machine that runs Excel.

Registry values (Office 2016 and later, including Microsoft 365, use version `16.0`):
- User setting: `HKCU\Software\Microsoft\Office\16.0\Excel\Security`, DWORD `AccessVBOM`, `1` = trusted.
- Group Policy: `HKCU\Software\Policies\Microsoft\Office\16.0\Excel\Security`, DWORD `AccessVBOM`.
  A policy value wins over the user value. Do not try to change it.

```python
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
```

The final proof is a real access. Do both checks:

```python
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
```

French UI path: Fichier > Options > Centre de gestion de la confidentialité >
Paramètres du Centre de gestion de la confidentialité > Paramètres des macros >
« Accès approuvé au modèle d'objet du projet VBA ».

PowerShell 5.1 check (read only):

```text
Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name AccessVBOM
```

Enable only with the developer's explicit approval (opt-in CLI flag, logged):

```text
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name AccessVBOM -Value 1 -Type DWord
```

Excel reads this value at start. Set it before `DispatchEx`.

---

## 6. Workbook and sheets

```python
def new_workbook(app):
    """Create a workbook with exactly one worksheet."""
    return app.Workbooks.Add(XL_WBAT_WORKSHEET)

def add_sheet_after_last(wb, name: str):
    """Add a worksheet at the end and name it."""
    last = wb.Worksheets(wb.Worksheets.Count)
    ws = wb.Worksheets.Add(None, last)        # positional: (Before, After). Tested.
    ws.Name = name
    return ws
```

Rules:
- **Tested: named arguments can be silently ignored with late binding.** `Worksheets.Add(After=last)`
  put the new sheet *before* the others, with no error. Always pass arguments **by position**.
  To skip an optional argument before the one you need, pass `None` (tested on `Worksheets.Add`).
  `pythoncom.Missing` did **not** work there. After any placement, check the sheet order in a test.
- COM collections are **1-based**: `wb.Worksheets(1)` is the first sheet.
- **Tested:** a new one-sheet workbook (`Workbooks.Add(-4167)`) is *named after its sheet*:
  `Feuil1` in French, `Sheet1` in English. `Application.Run("'Feuil1'!module.proc")` then fails.
  Save the workbook (`save_as_xlsm`) before any `Application.Run`.
- **Tested:** properties with parameters may come back as a whole tuple. `app.International(3)` raised
  `TypeError: 'tuple' object is not callable`; `app.International[2]` (0-based) worked.
- The first sheet is called `Feuil1` in French Excel and `Sheet1` in English. Rename it at once:
  `wb.Worksheets(1).Name = "Data"`.
- Tab names: max 31 characters, none of `: \ / ? * [ ]`, not empty, unique (case-insensitive).
- Hide a sheet: `ws.Visible = XL_SHEET_VERY_HIDDEN`.

---

## 7. Data, formulas, names

```python
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
```

```python
import datetime as dt

_EXCEL_EPOCH = dt.datetime(1899, 12, 30)

def to_excel_serial(value: dt.date | dt.datetime) -> float:
    """Pure: convert a date or naive datetime to an Excel serial number (no time zone shift)."""
    if not isinstance(value, dt.datetime):
        value = dt.datetime(value.year, value.month, value.day)
    if value.tzinfo is not None:
        raise ValueError("Pass a naive (wall-clock) datetime")
    return (value - _EXCEL_EPOCH).total_seconds() / 86400
```

Write it with `cell.Value2 = to_excel_serial(d)` (or inside a block), then set the format with
`put_en_us(cell, "NumberFormat", "yyyy-mm-dd")` (see 7.1 below).

Do not use `Range.Offset` for this. Its indexing through COM is easy to get wrong.
`ws.Cells(row, col)` is 1-based and always safe.

Rules (Documented):
- **Text is converted by Excel unless it starts with an apostrophe (tested, French Excel, from Python):**
  `"true"` -> boolean, `"1.10"` -> 1.1, `"2026-10-01"` and `"1/2"` -> dates, `"007"` -> 7, `"50%"` -> 0.5,
  `"#N/A"` -> an error value, `"=1+1"` -> a formula. With `as_cell()` all 13 tested texts stayed text, and
  `.Value` reads them back without the apostrophe. `write_block` applies it; do the same for any single write
  (`cell.Value = as_cell(text)`). Never apply it to formulas you want (`.Formula`) or to numbers and dates.
- All rows must have the same length. Pad with `None` (gives an empty cell).
- Pass numbers as `int`/`float`, never as strings like `"3,14"`.
- **Dates (tested):** pywin32 treats a **naive** `datetime` as local time and shifts it to UTC.
  In Paris, `datetime(2026, 11, 3)` was stored as 2 Nov 23:00. `datetime.date` raises `TypeError`.
  Safest: write the Excel serial number with `to_excel_serial()` (below) and set a `NumberFormat`.
  It does not depend on the time zone or the Excel language.
  Alternative: pass a UTC-aware datetime holding the wall-clock value you want
  (`datetime(2026, 11, 3, tzinfo=timezone.utc)`). It is stored exactly.
  Values read back are `pywintypes.datetime` tagged UTC. Use `.replace(tzinfo=None)` to compare.
- `Range.Formula` takes **English** function names and `,` separators: `"=SUM(A1:A10)"`.
  This works in French Excel too. `Range.FormulaLocal` is language-dependent: do not use it.
- **`NumberFormat` from Python is language-dependent (tested, French Excel).** pywin32 calls COM without
  the en-US locale, so French Excel reads the codes in French: `"0.00"` showed 5.5 as `0.06`,
  `"yyyy-mm-dd"` gave `yyyy-11-dd`, `"dd/mm/yyyy"` raised an error, `"0.0%"` showed 100 % as `10.0%`.
  Always set formats with `put_en_us()` below. It was tested with these five formats: all correct.
  VBA code is not affected: `.NumberFormat` in VBA always takes English codes (tested).
- `Range.Formula` from Python stays English (tested: `=ROUND(1.5,0)` gives 2 and shows `=ARRONDI(1,5;0)`).

### 7.1 Set locale-sensitive properties with the en-US LCID (tested)

```python
LCID_EN_US = 0x0409

def put_en_us(obj, prop: str, value) -> None:
    """Set a COM property with the en-US locale (LCID 1033) instead of pywin32's default."""
    ole = obj._oleobj_
    dispid = ole.GetIDsOfNames(prop)
    ole.Invoke(dispid, LCID_EN_US, pythoncom.DISPATCH_PROPERTYPUT, 0, value)
```

Use it for `NumberFormat` (and for any other property that takes format codes).
Read results with `.Value2`, never with `.Text`: `.Text` is the displayed string. It depends on the
format, the language and the column width.
- Defined name for VBA parameters:
  `wb.Names.Add("cfg_Region", "=Config!$B$2")`. VBA reads `ThisWorkbook.Names("cfg_Region").RefersToRange.Value`.

---

## 8. VBA components

### 8.1 Access

```python
def components(wb):
    """Return the VBComponents collection of a workbook."""
    return wb.VBProject.VBComponents
```

`VBComponent` members used here (Documented, VBA Extensibility 5.3):
- `.Name` (read/write), `.Type` (see constants), `.CodeModule`, `.Properties("Name").Value`.
- `VBComponents.Add(type)` with type 1, 2 or 3. Type 100 cannot be added.
- `VBComponents.Remove(component)` for types 1, 2, 3. Document modules cannot be removed.
- `VBComponents(name)` or `VBComponents.Item(name)` gets a component by name.

`CodeModule` members:
- `.CountOfLines`, `.Lines(start_line, count)` (1-based),
- `.DeleteLines(start_line, count)`, `.InsertLines(line, text)`, `.AddFromString(text)`.

### 8.2 Add or replace a standard module

```python
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
```

Why delete first: if the VBE option "Require Variable Declaration" is on, a new module already
contains `Option Explicit`. Adding source that also starts with `Option Explicit` gives a
duplicate line, which is a compile error. Deleting all lines first avoids this.

Module names: start with a letter, then letters, digits or `_`, max 31 characters,
unique in the project, not a VBA keyword. Use prefixes: `mod` (standard), `cls` (class).

### 8.3 Document modules (ThisWorkbook and sheets)

Event code (`Workbook_Open`, `Worksheet_Change`) goes in document modules.
Their names depend on the language and on the sheet. Find them without names:

```python
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
```

**Tested:** `ws.CodeName` is an empty string for a sheet just added by code, until the VBA project
has been accessed (`''` before, `'Feuil3'` after, in French Excel). Touching `wb.VBProject` first fixes it.
The fallback scan covers any other case. Default code names are `Feuil1`, `Feuil2`... in French
and `Sheet1`, `Sheet2`... in English, so always set your own (`shData`).

Set stable code names so VBA does not depend on the language:

```python
def set_sheet_code_name(wb, ws, code_name: str) -> None:
    """Give a sheet a language-independent code name, e.g. shData."""
    sheet_component(wb, ws).Name = code_name
```

After this, VBA can write `shData.Range("A1").Value`. This works in French and English Excel.

---

## 9. VBA source text rules

```python
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
```

Rules:
- VBA keywords are English in every Excel language. The source is the same for FR and EN.
- VBA stores code in the Windows ANSI code page (cp1252 on French and English Windows).
  Characters outside it become `?`. Keep source ASCII. Use `ChrW(233)` for `é` if needed.
- Do not include `Attribute VB_Name = ...` lines. They exist only in exported `.bas` files.
  `AddFromString` treats them as code and reports a syntax error.
- Read back code with `module.Lines(1, module.CountOfLines)`. VBE may change spacing or
  keyword case. To compare, normalize both sides, and compare case-insensitively if needed.

Minimal self-test module:

```python
MOD_SELF_TEST = r'''
Option Explicit

Public Function Ping() As String
    Ping = "pong"
End Function
'''
```

---

## 10. Buttons linked to macros

```python
def add_macro_button(ws, cell_address: str, caption: str, macro: str,
                     width: float = 120, height: float = 28):
    """Add a Form button at a cell that runs macro (e.g. 'modMain.Run')."""
    cell = ws.Range(cell_address)
    shape = ws.Shapes.AddFormControl(XL_BUTTON_CONTROL, cell.Left, cell.Top, width, height)
    shape.OnAction = macro
    shape.TextFrame.Characters().Text = caption   # Tested
    return shape
```

- `Shapes.AddFormControl(type, left, top, width, height)` is Documented.
- `Shape.OnAction` takes the macro name. Use `"moduleName.ProcName"` to avoid ambiguity.
- Tested: caption `'Lancer'` and `OnAction` `'modMain.Run'` read back unchanged, also after
  save and reopen. If a future Excel adds a workbook prefix (`'book.xlsm'!modMain.Run`),
  compare the part after `!`.

---

## 11. Save, close, reopen

```python
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
```

Rules:
- Always pass an **absolute** path. Excel's current folder is not Python's.
- Use positional arguments for `SaveAs`: `(Filename, FileFormat)`.
- After `SaveAs`, `wb.FileFormat` should be `52`. Check it in tests.
- Close with `wb.Close(False)` (do not save again).
- With macros disabled, the VBA project is still readable. `Application.Run` will not work.

---

## 12. Compile and run checks (integration tests only)

### 12.1 What proves the VBA is valid (tested)

- **`Ping` alone is not enough.** VBA compiles modules on demand. A file with a compile error in
  `modBad` still returned `"pong"` from `modSelfTest.Ping`.
- **Compile the whole project** with the VBE command "Debug > Compile VBAProject"
  (French: "Débogage > Compiler VBAProject"), control ID **578**. After a successful compile the
  command becomes disabled. If it stays enabled, the project has a compile error.
  Tested both ways in a hidden Excel: no dialog appeared, about 2 s plus Excel start.

```python
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
```

- `compile_vba_project` works with macros disabled (`AutomationSecurity` force disable). Tested.
- `run_macro` needs macros enabled: set `app.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW`
  before `Workbooks.Open`. Do this only for files your own script generated.
- Keep `EnableEvents = False` (set by `start_excel`) so `Workbook_Open` does not run during checks.

### 12.2 Always run checks in a child process with a timeout

A real file may still trigger a modal dialog (for example a runtime error in a macro you run).
A dialog in a hidden Excel blocks forever. So checks run in a child process. The child prints
its Excel PID on **stderr** first, and its result as the **last stdout line**.

```python
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

`subprocess.run` kills the child Python on timeout, but not the Excel it started.
`TimeoutExpired.stderr` is always `bytes`, even with `text=True`. That is why it is decoded.

Child script template (`_check_child.py`, next to the module that defines the helpers):

```python
"""Child: open an .xlsm, compile its VBA, optionally run Ping, print the result."""
import sys
from pathlib import Path

from xlsm_builder.excel_checks import (   # adapt to the module that holds the helpers
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

Call it with `run_child(script, path)` or `run_child(script, path, "--ping")`.
Expected last lines: `COMPILED_OK`, `COMPILED_OK pong`, or `COMPILE_FAILED`.

## 13. Delivery to other machines

The generating machine is not the user's machine. Plan this:

- Files sent by email or downloaded get the **Mark of the Web**. Since 2022, Office on Windows
  blocks their macros with no "Enable content" button.
- Ways to open them with macros:
  - the recipient unblocks the file: right-click > Properties > General > check **Unblock**;
  - PowerShell: `Unblock-File -Path "C:\path\file.xlsm"`;
  - Python: `os.remove(r"C:\path\file.xlsm:Zone.Identifier")` (deletes the mark stream);
  - a Trusted Location set by IT;
  - files synced through OneDrive are not marked.
- Each user must also allow macros (Trust Center) or IT must sign and trust the publisher.
- VBA does not run in Excel for the web or on mobile.

---

## 14. Not possible with this stack

- Signing a VBA project from code. pywin32 has no API for it. Signing is manual in the VBE
  (Tools > Digital Signature), and any later code change removes the signature.
- Running without Excel installed. COM needs Excel desktop on the generating machine.
- Generating on Linux or in a Linux container.

---

## 15. Fakes for unit tests (pattern)

```python
class FakeCodeModule:
    def __init__(self) -> None:
        self.text = ""
    @property
    def CountOfLines(self) -> int:
        return len(self.text.splitlines())
    def DeleteLines(self, start: int, count: int) -> None:
        lines = self.text.splitlines(keepends=True)
        del lines[start - 1:start - 1 + count]
        self.text = "".join(lines)
    def AddFromString(self, code: str) -> None:
        self.text += code

class FakeComponent:
    def __init__(self, name: str, comp_type: int) -> None:
        self.Name = name
        self.Type = comp_type
        self.CodeModule = FakeCodeModule()

class FakeComponents:
    def __init__(self) -> None:
        self.items: list[FakeComponent] = []
    @property
    def Count(self) -> int:
        return len(self.items)
    def Item(self, index):
        if isinstance(index, int):
            return self.items[index - 1]
        return next(c for c in self.items if c.Name == index)
    def Add(self, comp_type: int) -> FakeComponent:
        comp = FakeComponent(f"Module{len(self.items) + 1}", comp_type)
        self.items.append(comp)
        return comp

class FakeVBProject:
    def __init__(self) -> None:
        self.VBComponents = FakeComponents()

class FakeWorkbook:
    def __init__(self) -> None:
        self.VBProject = FakeVBProject()
        self.CodeName = "ThisWorkbook"
```

Example unit test:

```python
def test_add_or_replace_module_replaces_code():
    wb = FakeWorkbook()
    add_or_replace_module(wb, "modMain", "Sub A()\r\nEnd Sub\r\n")
    add_or_replace_module(wb, "modMain", "Sub B()\r\nEnd Sub\r\n")
    comp = wb.VBProject.VBComponents.Item("modMain")
    assert comp.CodeModule.text == "Sub B()\r\nEnd Sub\r\n"
    assert wb.VBProject.VBComponents.Count == 1
```

Keep fakes small. Add only the members the tested function uses.
