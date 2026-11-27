Attribute VB_Name = "modRovoDebug"
Option Explicit
' The button VBA code for Rovo (AL, 2026-11-27: "Rovo cannot read the macros of an xlsm"): ONE text file with the code of the delivered modules and
' the formula of every column of the tables, for Rovo to read the code behind a figure. The code is the hidden sheet VbaCode (Module | Line | Text),
' written at build from vba\*.bas (src/jira2dashboard/vba_code.py): the workbook never reads its own VBA project (that needs "Trust access to the VBA
' project object model", and an endpoint protection reads a macro that reads macros as malware). The formulas are the sheet FlowFormulas. The file
' vba-code.txt goes into the folder rovo-debug next to the workbook, UTF-8 without BOM (WriteUtf8File: plain VBA, no ADODB.Stream), lines ended by LF.
' It replaces the files 01.md, 02.md... of the button Debug prompt for Rovo (v0.20.0), deleted at each click. The log gets counts only, no path.
' Twin: vba_code.file_text (the Excel scenario compares the file with it, character by character).

Public Const SHEET_VBA_CODE As String = "VbaCode"
Public Const ROVO_DEBUG_FOLDER As String = "rovo-debug"
Public Const VBA_CODE_FILE As String = "vba-code.txt"
Private Const TITLE As String = "# jira2dashboard - the code of the workbook, for Rovo"
Private Const INTRO As String = "This file holds the VBA modules of the workbook (the macros of the buttons) and the formula of every column of its tables (sheet FlowFormulas_from_excel). It holds no value of Jira and no setting. Read the code to explain a figure; quote the module and the line."
Private Const FORMULAS_HEAD As String = "===== Formulas of the tables (sheet FlowFormulas_from_excel): Table | Column | Formula ====="

Public Function RovoDebugFolder(ByVal workbookPath As String, ByVal localAppData As String, ByVal temp As String) As String
    ' rovo-debug next to the workbook; under the folder of the user when the workbook has no folder of its own (never saved, or opened from
    ' SharePoint: an http path), as modLog.LogFolderFrom.
    Cov "RVD01"
    If Len(workbookPath) > 0 And LCase$(Left$(workbookPath, 4)) <> "http" Then RovoDebugFolder = workbookPath & "\" & ROVO_DEBUG_FOLDER: Exit Function
    Cov "RVD02"
    If Len(localAppData) = 0 Then localAppData = temp
    RovoDebugFolder = localAppData & "\jira2dashboard\" & ROVO_DEBUG_FOLDER
End Function

Public Function WriteVbaCode(ByVal wb As Object, ByVal folder As String) As String
    ' Writes vba-code.txt into `folder` (the 01.md .. 99.md of the earlier button deleted) and returns "modules|lines|characters|formulas". Error 635
    ' (J2D-ROVO-006) when the sheet VbaCode is missing or holds no code.
    Cov "RVD03"
    Dim code As Variant, formulas As Variant, nCode As Long, nFormulas As Long, out() As String, i As Long, k As Long, list As String
    Dim names As New Collection, counts As Object, nm As Variant, current As String, text As String, n As Long
    code = SheetBlock(wb, SHEET_VBA_CODE, nCode)
    If nCode = 0 Then
        Cov "RVD04"
        Err.Raise vbObjectError + 635, "WriteVbaCode", "The sheet " & SheetName(SHEET_VBA_CODE) & " holds no code: use a fresh copy of the delivered workbook."
    End If
    formulas = SheetBlock(wb, "FlowFormulas", nFormulas)
    Set counts = CreateObject("Scripting.Dictionary")
    For i = 1 To nCode
        nm = CStr(code(i, 1) & "")
        If Not counts.Exists(nm) Then counts(nm) = 0: names.Add nm
        counts(nm) = counts(nm) + 1
    Next i
    For Each nm In names
        list = list & IIf(Len(list) > 0, ", ", "") & nm & " (" & counts(nm) & " lines)"
    Next nm
    ReDim out(1 To 4 + 2 * names.Count + nCode + 2 + nFormulas)          ' the lines in an array, joined once: a long text made by & would cost minutes
    AddLine out, k, TITLE
    AddLine out, k, "Version: " & WorkbookVersion(wb)
    AddLine out, k, INTRO
    AddLine out, k, "Modules: " & list
    For i = 1 To nCode
        If CStr(code(i, 1) & "") <> current Then
            current = CStr(code(i, 1) & "")
            AddLine out, k, ""
            AddLine out, k, "===== VBA module " & current & " (" & counts(current) & " lines) ====="
        End If
        AddLine out, k, CellString(code(i, 3))
    Next i
    AddLine out, k, ""
    AddLine out, k, FORMULAS_HEAD
    For i = 1 To nFormulas
        AddLine out, k, CellString(formulas(i, 1)) & " | " & CellString(formulas(i, 2)) & " | " & CellString(formulas(i, 3))
    Next i
    text = Join(out, vbLf) & vbLf
    MakeFolder folder
    On Error Resume Next
    For n = 1 To 99                                                     ' the files of the earlier button only (a wildcard ??.md also takes x.md)
        Kill folder & "\" & Format$(n, "00") & ".md"
    Next n
    Err.Clear
    On Error GoTo 0
    WriteUtf8File folder & "\" & VBA_CODE_FILE, text
    WriteVbaCode = names.Count & "|" & nCode & "|" & Len(text) & "|" & nFormulas
End Function

Private Sub AddLine(ByRef out() As String, ByRef k As Long, ByVal text As String)
    k = k + 1
    out(k) = text
End Sub

Private Function CellString(ByVal v As Variant) As String
    ' The text of a cell: "" for an empty one or an error (a line of code is never an error: written as text at build).
    If IsError(v) Then Exit Function
    CellString = CStr(v & "")
End Function

Private Function SheetBlock(ByVal wb As Object, ByVal name As String, ByRef n As Long) As Variant
    ' The cells of columns A to C under the header of a sheet, in one read; n = their rows (column A is filled on every row). 0 without the sheet.
    Dim ws As Object, last As Long
    n = 0
    If Not SheetExists(wb, name) Then Exit Function
    Set ws = wb.Worksheets(SheetName(name))
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row                    ' xlUp
    If last < 2 Then Exit Function
    n = last - 1
    SheetBlock = ws.Range(ws.Cells(2, 1), ws.Cells(last, 3)).Value       ' 3 columns: always an array, even for one row
End Function

Public Sub WriteUtf8File(ByVal path As String, ByVal text As String)
    ' The text as UTF-8 without BOM (modText.Utf8Encode: plain VBA, no ADODB.Stream). An older file is deleted first: Put into a Binary file never
    ' makes it shorter, the end of a longer older file would stay.
    Cov "RVD05"
    Dim f As Integer, bytes() As Byte, num As Long, desc As String
    On Error Resume Next
    Kill path
    Err.Clear
    On Error GoTo Failed
    f = FreeFile
    Open path For Binary Access Write As #f
    If Len(text) > 0 Then bytes = Utf8Encode(text): Put #f, , bytes
    Close #f
    Exit Sub
Failed:
    Cov "RVD06"
    num = Err.Number: desc = Err.Description                            ' first: Close may clear Err
    On Error Resume Next
    Close #f                                                            ' a file left open would refuse the next click
    On Error GoTo 0
    Err.Raise num, "WriteUtf8File", desc
End Sub

Private Sub MakeFolder(ByVal folder As String)
    ' Each missing folder of the path (MkDir fails harmlessly on one that exists).
    Dim parts() As String, i As Long, path As String
    On Error Resume Next
    parts = Split(folder, "\")
    path = parts(0)
    For i = 1 To UBound(parts)
        path = path & "\" & parts(i)
        MkDir path
    Next i
    Err.Clear
End Sub
