"""Probe: NumberFormat from Python (default call vs forced en-US LCID) and from VBA, in French Excel."""
import datetime as dt
import pythoncom
import xlsm_ref as R

LCID_EN_US = 0x0409

VBA = r'''
Option Explicit
Public Function SetFmtFromVba(ByVal addr As String, ByVal fmt As String) As String
    With ThisWorkbook.Worksheets(1).Range(addr)
        .NumberFormat = fmt
        SetFmtFromVba = .Text & " | local=" & .NumberFormatLocal
    End With
End Function
'''


def put_en_us(obj, prop: str, value) -> None:
    """Set a COM property with the en-US locale (LCID 1033) instead of pywin32's default."""
    ole = obj._oleobj_
    dispid = ole.GetIDsOfNames(prop)
    ole.Invoke(dispid, LCID_EN_US, pythoncom.DISPATCH_PROPERTYPUT, 0, value)


def try_set(c, fmt, how):
    try:
        if how == "default":
            c.NumberFormat = fmt
        else:
            put_en_us(c, "NumberFormat", fmt)
        return f"Text={c.Text!r:18} Local={c.NumberFormatLocal!r}"
    except Exception as e:
        return f"ERROR {type(e).__name__}"


DAY = R.to_excel_serial(dt.date(2026, 11, 3))
CASES = [(1.0, "0.0%"), (5.5, "0.00"), (1234567.891, "#,##0.00"), (DAY, "yyyy-mm-dd"), (DAY, "dd/mm/yyyy")]

with R.excel_session() as app:
    wb = R.new_workbook(app)
    ws = wb.Worksheets(1)
    print("decimal separator:", app.International[2])
    for how in ("default", "en-US LCID"):
        print(f"--- Python, {how} ---")
        for i, (value, fmt) in enumerate(CASES, start=1):
            c = ws.Cells(i, 1 if how == "default" else 2)
            c.Value2 = value
            print(f"{fmt:12} -> {try_set(c, fmt, how)}")
    print("--- VBA ---")
    R.add_or_replace_module(wb, "modProbe", R.normalize_vba(VBA))
    print("unsaved workbook name:", wb.Name)
    from pathlib import Path
    R.save_as_xlsm(wb, Path(__file__).parent / "out" / "probe_numfmt.xlsm", overwrite=True)
    for i, (value, fmt) in enumerate(CASES, start=1):
        ws.Cells(i, 3).Value2 = value
        print(f"{fmt:12} -> {app.Run(f'{chr(39)}{wb.Name}{chr(39)}!modProbe.SetFmtFromVba', f'C{i}', fmt)}")
    print("--- Formula check, default call ---")
    ws.Range("D1").Formula = "=ROUND(1.5,0)"
    print("=ROUND(1.5,0) ->", ws.Range("D1").Value2, ws.Range("D1").FormulaLocal)
    wb.Close(False)
    del c, ws, wb
