"""Patch 7: every text written to a cell gets a leading apostrophe (Excel converts true, 1.10, dates, 007, 50%...)."""
from pathlib import Path

D = Path(__file__).parent


def rep(name, old, new):
    p = D / name
    s = p.read_text(encoding="ascii").replace("\r\n", "\n")
    assert s.count(old) == 1, (name, old[:60])
    p.write_text(s.replace(old, new), encoding="ascii", newline="\n")


rep("modJira.bas", """Public Function SafeCell(ByVal value As Variant) As Variant
    ' Text starting with = + - @ would become a formula when written with Range.Value.
    If VarType(value) = vbString Then
        If Len(value) > 0 Then
            If InStr("=+-@", Left$(value, 1)) > 0 Then Cov "JIR12": SafeCell = "'" & value: Exit Function
        End If
    End If
    SafeCell = value
End Function""", """Public Function SafeCell(ByVal value As Variant) As Variant
    ' Every non-empty text gets a leading apostrophe, so Excel keeps it as text (tested).
    ' Without it Excel converts: "true" -> TRUE, "1.10" -> 1.1, "2026-10-01" and "1/2" -> dates,
    ' "007" -> 7, "50%" -> 0.5, "#N/A" -> an error value, "=1+1" -> a formula.
    ' The apostrophe is not part of the value: .Value reads the text back unchanged.
    If VarType(value) = vbString Then
        If Len(value) > 0 Then Cov "JIR12": SafeCell = "'" & Left$(value, 32766): Exit Function
    End If
    SafeCell = value
End Function""")

rep("modTests.bas", """    Check "sheet", "normal text unchanged", "abc", "abc", SafeCell("abc")""",
    """    Check "sheet", "any text gets the apostrophe", "abc", "'abc", SafeCell("abc")
    Check "sheet", "empty text unchanged", "(empty)", "", SafeCell("")
    Check "sheet", "long text cut to fit a cell", "40000 x", 32767, Len(SafeCell(String$(40000, "x")))""")

rep("modTests.bas", """        "T_TooManyPages", "T_BackgroundExport")""",
    """        "T_TooManyPages", "T_BackgroundExport", "T_TextNotConverted")""")

rep("modTests.bas", """Public Sub T_StringifyBadObject()""", """Public Sub T_TextNotConverted()
    ' Texts that Excel would turn into booleans, numbers, dates, errors or formulas stay text.
    Dim ws As Object, rows As New Collection, texts As Variant, c As Long
    texts = Array("true", "TRUE", "1.10", "2026-10-01", "1/2", "007", "50%", "#N/A", "=1+1", "-x", "PI3-S1")
    Set ws = ThisWorkbook.Worksheets("TestScratch")
    rows.Add texts
    WriteTable ws, texts, rows
    For c = 0 To UBound(texts)
        Check "sheet", "text stays text", CStr(texts(c)), CStr(texts(c)) & "|String", _
              CStr(ws.Cells(2, c + 1).Value) & "|" & TypeName(ws.Cells(2, c + 1).Value)
    Next c
End Sub

Public Sub T_StringifyBadObject()""")
print("patched")
