Attribute VB_Name = "modParams"
Option Explicit
' Parameters live in the workbook, never in code. The sheets Workflows and Transitions are written by
' the Python generator from params\*.md; this module only reads them. Each reader takes the sheet as an argument
' (testable on a scratch sheet). Every call re-reads the sheet: keep the result (for example the Dictionary of
' StatusCategories) when you need it in a loop.

Public Const SHEET_WORKFLOWS As String = "Workflows"
Public Const SHEET_TRANSITIONS As String = "Transitions"
Public Const SHEET_ROVO_SPEC As String = "RovoPromptSpec"
Public Const SHEET_ROVO As String = "RovoPrompt"
' The warnings of RuleWarnings: the same texts as flow/status_rules.py (tests/unit/test_status_rules.py reads them here).
Private Const NO_START_TEXT As String = " workflow: Flow Time, Flow Efficiency and Flow Load cannot be computed | declare IsStart = 1 for its first in-progress status in the sheet Workflows"
Private Const NO_WAIT_TEXT As String = "no status is declared Wait: Flow Efficiency cannot be computed | declare Category = Wait in the sheet Workflows for the statuses where the work waits"

Public Function ReadTable(ByVal ws As Object) As Collection
    ' Row 1 holds the headers. Returns one Dictionary (header -> cell value) per row whose first cell is not empty.
    Cov "PRM01"
    Dim data As Variant, out As New Collection, r As Long, c As Long, d As Object, nRows As Long, nCols As Long
    Set ReadTable = out
    nRows = ws.Cells(ws.Rows.Count, 1).End(-4162).Row            ' -4162 = xlUp
    nCols = ws.Cells(1, ws.Columns.Count).End(-4159).Column      ' -4159 = xlToLeft
    If nRows < 2 Then Cov "PRM02": Exit Function
    data = ws.Range(ws.Cells(1, 1), ws.Cells(nRows, nCols)).Value
    For r = 2 To nRows
        If Len(CStr(data(r, 1) & "")) > 0 Then
            Set d = CreateObject("Scripting.Dictionary")
            For c = 1 To nCols
                d(CStr(data(1, c))) = data(r, c)
            Next c
            out.Add d
        End If
    Next r
End Function

Public Function StatusCategories(ByVal wsWorkflows As Object, ByVal workflow As String) As Object
    ' status -> category ("Active", "Wait", ...) for one workflow. This is the input of BuildStatusIntervals.
    Cov "PRM03"
    Dim out As Object, row As Variant
    Set out = CreateObject("Scripting.Dictionary")
    For Each row In ReadTable(wsWorkflows)
        If CStr(row("Workflow")) = workflow Then out(CStr(row("Status"))) = CStr(row("Category"))
    Next row
    Set StatusCategories = out
End Function

Public Function StatusHasFlag(ByVal wsWorkflows As Object, ByVal workflow As String, ByVal status As String, _
                              ByVal flagColumn As String) As Boolean
    Cov "PRM06"
    Dim row As Variant
    For Each row In ReadTable(wsWorkflows)
        If CStr(row("Workflow")) = workflow Then
            If CStr(row("Status")) = status Then StatusHasFlag = (Val(CStr(row(flagColumn) & "")) = 1): Exit Function
        End If
    Next row
End Function

Public Function ProjectKeys(ByVal wsProjects As Object) As Collection
    ' Keys of the Projects sheet (columns Project | Role | Team), the Program first.
    Cov "PRM18"
    Dim out As New Collection, rows As Collection, row As Variant
    Set rows = ReadTable(wsProjects)
    For Each row In rows
        If StrComp(TrimWhite(CStr(row("Role") & "")), "Program", vbTextCompare) = 0 Then out.Add TrimWhite(CStr(row("Project")))
    Next row
    For Each row In rows
        If StrComp(TrimWhite(CStr(row("Role") & "")), "Program", vbTextCompare) <> 0 Then out.Add TrimWhite(CStr(row("Project")))
    Next row
    Set ProjectKeys = out
End Function

Public Function ProgramProject(ByVal wsProjects As Object) As String
    ' The project whose Role is Program (the train, it holds the Features). Empty text if the sheet has none.
    Cov "PRM19"
    Dim row As Variant
    For Each row In ReadTable(wsProjects)
        If StrComp(TrimWhite(CStr(row("Role") & "")), "Program", vbTextCompare) = 0 Then
            ProgramProject = TrimWhite(CStr(row("Project"))): Exit Function
        End If
    Next row
End Function

Public Function ParseMoment(ByVal text As String) As Variant
    ' "2026-11-03", "2026-11-03 18:00" or "2026-11-03T18:00:00" (UTC) -> Date. Empty text if blank or invalid.
    ' Parsed by hand: CDate depends on the Excel language.
    Cov "PRM21"
    Dim s As String, d As Date, parts() As String, h As Long, m As Long, sec As Long
    ParseMoment = Empty
    s = TrimWhite(text)
    If UCase$(Right$(s, 1)) = "Z" Then s = Left$(s, Len(s) - 1)         ' UTC said in so many words
    If Len(s) < 10 Or Len(s) > 19 Then Exit Function                  ' longer: an offset (+01:00) or milliseconds, refused as the Python path refuses them
    If Mid$(s, 5, 1) <> "-" Or Mid$(s, 8, 1) <> "-" Then Exit Function
    On Error GoTo Invalid
    d = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2)))
    If Month(d) <> CInt(Mid$(s, 6, 2)) Or Day(d) <> CInt(Mid$(s, 9, 2)) Then GoTo Invalid     ' 2026-13-45 does not exist
    If Len(s) > 10 Then                                              ' a time: H:MM or H:MM:SS, as TIMEVALUE(MID(t,12,8)) of the sheet reads it
        parts = Split(Mid$(s, 12, 8), ":")
        If UBound(parts) < 1 Or UBound(parts) > 2 Then GoTo Invalid  ' "2026-11-03 18" is no time; it was dropped without a word
        h = CInt(parts(0)): m = CInt(parts(1))
        If UBound(parts) = 2 Then sec = CInt(parts(2))
        If h > 23 Or m > 59 Or sec > 59 Or h < 0 Or m < 0 Or sec < 0 Then GoTo Invalid
        d = d + TimeSerial(h, m, sec)
    End If
    ParseMoment = d
    Exit Function
Invalid:
    Err.Clear
    ParseMoment = Empty
End Function

Public Function AsOfFromConfig(ByVal wsConfig As Object) As Date
    ' The AsOf parameter of the Config sheet (UTC). The metrics need it: open statuses run up to it.
    Cov "PRM22"
    Dim v As Variant, row As Variant
    For Each row In ReadTable(wsConfig)
        If CStr(row("Parameter")) = "AsOf" Then v = row("Value"): Exit For
    Next row
    If VarType(v) = vbDate Or VarType(v) = vbDouble Then
        Cov "PRM24"                                  ' typed as a date: Excel keeps the date, not the text
        If CDbl(v) < 36526 Or CDbl(v) >= 73051 Then  ' 2000 to 2099: a plain number such as 5 or 20261120 was typed, not a date
            Cov "PRM25"
            Err.Raise vbObjectError + 604, "AsOfFromConfig", "AsOf in the Config sheet is not a date (" & CStr(v) & "): type a UTC date and time such as 2026-11-03 18:00"
        End If
        AsOfFromConfig = CDate(v): Exit Function
    End If
    v = ParseMoment(CStr(v & ""))
    If IsEmpty(v) Then
        Cov "PRM23"
        Err.Raise vbObjectError + 604, "AsOfFromConfig", "Type AsOf in the Config sheet: a UTC date and time such as 2026-11-03 18:00"
    End If
    If CDbl(v) < 36526 Or CDbl(v) >= 73051 Then
        Cov "PRM73"
        Err.Raise vbObjectError + 604, "AsOfFromConfig", "AsOf in the Config sheet must be from 2000 to 2099: type a UTC date and time such as 2026-11-03 18:00"
    End If
    AsOfFromConfig = v
End Function

Public Function ConfigValue(ByVal wsConfig As Object, ByVal name As String, Optional ByVal defaultValue As String = "") As String
    ' One value of the Config sheet (columns read by header name: Parameter, Value), trimmed like a pasted value.
    ' An empty cell or an unknown parameter gives defaultValue.
    Cov "PRM14"
    Dim row As Variant, v As String
    For Each row In ReadTable(wsConfig)
        If CStr(row("Parameter")) = name Then
            v = TrimWhite(CStr(row("Value") & ""))
            If Len(v) = 0 Then Cov "PRM15": v = defaultValue
            ConfigValue = v
            Exit Function
        End If
    Next row
    Cov "PRM16"
    ConfigValue = defaultValue
End Function

Public Function MissingMandatory(ByVal wsConfig As Object, Optional ByVal onlyNames As String = "") As String
    ' The parameters of the Config sheet whose Mandatory is True and whose Value is empty, as "BaseUrl, AsOf" ("" = none),
    ' in the order of the sheet. onlyNames ("BaseUrl,Email") limits the check to those parameters; "" = all of them.
    ' A sheet without the column Mandatory (a workbook built before it) has nothing to check.
    Cov "PRM26"
    Dim row As Variant, name As String, out As String
    For Each row In ReadTable(wsConfig)
        If Not row.Exists("Mandatory") Then Cov "PRM27": Exit Function
        name = CStr(row("Parameter") & "")
        If IsYes(CStr(row("Mandatory") & "")) And Len(TrimWhite(CStr(row("Value") & ""))) = 0 Then
            If Len(onlyNames) = 0 Or InStr(1, "," & Replace(onlyNames, " ", "") & ",", "," & name & ",", vbTextCompare) > 0 Then
                Cov "PRM28"
                If Len(out) > 0 Then out = out & ", "
                out = out & name
            End If
        End If
    Next row
    MissingMandatory = out
End Function

Public Sub RequireMandatory(ByVal wsConfig As Object, Optional ByVal onlyNames As String = "")
    ' Stops a button before any call to Jira while a mandatory value of the Config sheet is empty (error 608).
    Dim missing As String
    missing = MissingMandatory(wsConfig, onlyNames)
    If Len(missing) > 0 Then
        Cov "PRM29"
        Err.Raise vbObjectError + 608, "RequireMandatory", "Fill these values of the Config sheet first (Mandatory = True): " & missing
    End If
End Sub

Public Function LockReason(ByVal isReadOnly As Boolean, ByVal inProtectedView As Boolean) As String
    ' Why the workbook cannot be used, "" when it can: a workbook opened in Protected View (file from the Internet or a mail) or
    ' read-only. Pure: the state is read by RequireEditable.
    If inProtectedView Then
        Cov "PRM54": LockReason = "it is open in Protected View"
    ElseIf isReadOnly Then
        Cov "PRM55": LockReason = "it is open read-only"
    End If
End Function

Public Sub RequireEditableIf(ByVal isReadOnly As Boolean, ByVal inProtectedView As Boolean)
    ' Error 612 (J2D-SYS-003) when the workbook is locked: the message says what to do, before any call to Jira.
    Dim why As String
    why = LockReason(isReadOnly, inProtectedView)
    If Len(why) > 0 Then
        Cov "PRM10"
        Err.Raise vbObjectError + 612, "RequireEditable", "The workbook is locked: " & why & ". If a yellow bar is shown at the top of Excel, " & _
                  "click Enable Editing. Otherwise close what holds the file (another Excel or user), or save a copy (File > Save As) and open " & _
                  "that one; for a file that came from the Internet also right-click it > Properties > Unblock. Then click the button again."
    End If
End Sub

Public Function SameFile(ByVal pathA As String, ByVal pathB As String) As Boolean
    ' Windows paths are not case-sensitive.
    SameFile = (StrComp(pathA, pathB, vbTextCompare) = 0)
End Function

Public Sub RequireEditable()
    ' Stops a button that writes the workbook while it is locked. The probe never raises by itself (a failed probe = not locked).
    Dim views As Object, i As Long, inView As Boolean, ro As Boolean
    On Error Resume Next
    Set views = Application.ProtectedViewWindows
    If Not views Is Nothing Then
        For i = 1 To views.Count
            If SameFile(CStr(views(i).SourcePath) & "\" & CStr(views(i).SourceName), ThisWorkbook.FullName) Then inView = True
        Next i
    End If
    ro = ThisWorkbook.ReadOnly
    Err.Clear
    On Error GoTo 0
    RequireEditableIf ro, inView
End Sub

Public Function RovoPromptLines(ByVal wb As Object) As Collection
    ' The prompt for Rovo, one line per item: the rows of the sheet RovoPromptSpec (Part | Name | Text) in order. A text row
    ' is a line as it is, "config" becomes one line per parameter of the sheet Config, "sheet" one line for a parameter sheet.
    ' The Value column of Config is never read: it may hold the address, the e-mail or the token.
    ' The parameters of the parts "keep" (the connection) and "skip" (what the workbook reads in Jira) are not listed (AL, 2026-11-23).
    Cov "PRM30"
    Dim out As New Collection, spec As Variant, specs As Collection, unasked As Object
    Set specs = ReadTable(wb.Worksheets(SheetName(SHEET_ROVO_SPEC)))
    Set unasked = CreateObject("Scripting.Dictionary")
    For Each spec In specs
        Select Case LCase$(CStr(spec("Part") & ""))
            Case "keep", "skip": unasked(LCase$(CStr(spec("Name") & ""))) = True
        End Select
    Next spec
    For Each spec In specs
        Select Case LCase$(CStr(spec("Part") & ""))
            Case "keep", "skip", "load"                                  ' a parameter or a sheet that Rovo is not asked for: no line
            Case "config": Cov "PRM31": AddConfigLines wb.Worksheets(SheetName("Config")), out, unasked
            Case "sheet": Cov "PRM32": out.Add SheetLine(wb, CStr(spec("Name") & ""), CStr(spec("Text") & ""))
            Case Else: Cov "PRM33": out.Add CStr(spec("Text") & "")
        End Select
    Next spec
    Set RovoPromptLines = out
End Function

Private Sub AddConfigLines(ByVal wsConfig As Object, ByVal out As Collection, ByVal unasked As Object)
    ' "- Name (mandatory): meaning. Example: form of the value", from the columns Parameter, Mandatory, Description, Example; not for the parameters
    ' that Rovo is not asked for (unasked: lower-case name -> True).
    Dim row As Variant, text As String, example As String
    For Each row In ReadTable(wsConfig)
        If unasked.Exists(LCase$(CStr(row("Parameter") & ""))) Then Cov "PRM74": GoTo NextRow
        text = "- " & CStr(row("Parameter") & "") & IIf(IsYes(CStr(row("Mandatory") & "")), " (mandatory): ", " (optional): ") & _
               CStr(row("Description") & "")
        example = CStr(row("Example") & "")
        If Len(example) > 0 Then text = text & " Example: " & example
        out.Add text
NextRow:
    Next row
End Sub

Private Function SheetLine(ByVal wb As Object, ByVal name As String, ByVal rule As String) As String
    ' "SHEET Projects - columns: Project | Role | Team. rule", the columns read from the header row of the sheet.
    Dim ws As Object, c As Long, cols As String, head As String
    If Not SheetExists(wb, name) Then SheetLine = "SHEET " & name & " - this sheet is not in the workbook. " & rule: Exit Function
    Set ws = wb.Worksheets(SheetName(name))
    For c = 1 To ws.Columns.Count
        head = CStr(ws.Cells(1, c).Value & "")
        If Len(head) = 0 Then Exit For
        If head <> "Example" Then cols = cols & IIf(Len(cols) > 0, " | ", "") & head
    Next c
    SheetLine = "SHEET " & name & " - columns: " & cols & ". " & rule
End Function

Public Function WriteRovoPrompt(ByVal wb As Object) As Long
    ' Writes the prompt into the sheet RovoPrompt (created at the end when missing), column A, one line per row, and returns
    ' the number of lines. The user selects them, copies and pastes them into Rovo.
    Cov "PRM34"
    Dim lines As Collection, data() As Variant, i As Long, ws As Object
    Set lines = RovoPromptLines(wb)
    If lines.Count = 0 Then Err.Raise vbObjectError + 630, "WriteRovoPrompt", "The sheet " & SHEET_ROVO_SPEC & " has no row: rebuild the workbook."
    If SheetExists(wb, SHEET_ROVO) Then
        Set ws = wb.Worksheets(SheetName(SHEET_ROVO))
    Else
        Cov "PRM35"
        Set ws = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SheetName(SHEET_ROVO)
    End If
    ws.Cells.ClearContents
    ReDim data(1 To lines.Count, 1 To 1)
    For i = 1 To lines.Count
        data(i, 1) = SafeCell(CStr(lines(i)))
    Next i
    ws.Range(ws.Cells(1, 1), ws.Cells(lines.Count, 1)).Value = data
    ws.Columns(1).ColumnWidth = 150
    WriteRovoPrompt = lines.Count
End Function

' ---- The answer of Rovo (CSV: Sheet, Row, Column, Value, Note) loaded into the parameter sheets ---------------------------------

Public Function ParseCsv(ByVal text As String) As Collection
    ' CSV as Rovo or Excel export it (RFC 4180): the delimiter is the comma, the semicolon or the tab (the most frequent one on
    ' the first line outside quotes); a value in double quotes may hold the delimiter, a line break and "" for a quote. One
    ' array of strings per record; an empty line is no record.
    Cov "PRM36"
    Dim out As New Collection, delim As String, i As Long, n As Long, ch As String, inQ As Boolean, field As String
    Dim rec() As String, count As Long, started As Boolean
    delim = CsvDelimiter(text)
    n = Len(text)
    ReDim rec(0 To 15)
    For i = 1 To n
        ch = Mid$(text, i, 1)
        If inQ Then
            If ch = """" Then
                If Mid$(text, i + 1, 1) = """" Then field = field & """": i = i + 1 Else inQ = False
            Else
                field = field & ch
            End If
        ElseIf ch = """" And Len(field) = 0 Then
            inQ = True: started = True
        ElseIf ch = delim Then
            CsvAddField rec, count, field
            started = True
        ElseIf ch = vbCr Or ch = vbLf Then
            If ch = vbCr And Mid$(text, i + 1, 1) = vbLf Then i = i + 1
            If started Or Len(field) > 0 Then CsvAddField rec, count, field: CsvEmit out, rec, count
            started = False
        Else
            field = field & ch
        End If
    Next i
    If started Or Len(field) > 0 Then Cov "PRM37": CsvAddField rec, count, field: CsvEmit out, rec, count
    Set ParseCsv = out
End Function

Private Sub CsvAddField(ByRef rec() As String, ByRef count As Long, ByRef field As String)
    If count > UBound(rec) Then ReDim Preserve rec(0 To 2 * count + 1)
    rec(count) = field
    count = count + 1
    field = ""
End Sub

Private Sub CsvEmit(ByVal out As Collection, ByRef rec() As String, ByRef count As Long)
    Dim one() As String, i As Long
    ReDim one(0 To count - 1)
    For i = 0 To count - 1
        one(i) = rec(i)
    Next i
    out.Add one
    count = 0
End Sub

Private Function CsvDelimiter(ByVal text As String) As String
    ' The comma unless the semicolon or the tab is more frequent on the first line (outside quotes).
    Dim i As Long, ch As String, inQ As Boolean, commas As Long, semis As Long, tabs As Long
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch = """" Then
            inQ = Not inQ
        ElseIf Not inQ Then
            If ch = vbCr Or ch = vbLf Then Exit For
            If ch = "," Then commas = commas + 1
            If ch = ";" Then semis = semis + 1
            If ch = vbTab Then tabs = tabs + 1
        End If
    Next i
    CsvDelimiter = ","
    If semis > commas And semis >= tabs Then Cov "PRM38": CsvDelimiter = ";"
    If tabs > commas And tabs > semis Then Cov "PRM39": CsvDelimiter = vbTab
End Function

Public Function RovoCsvRead(ByVal wb As Object, ByVal text As String) As Object
    ' The CSV as a Dictionary: items = Collection of Array(sheet, row, column, value), empty = how many values are empty,
    ' unknown = how many values Rovo could not give (UNKNOWN, or the token that the user types), problems = Collection of texts. Only the sheets of the prompt are
    ' accepted (Config and the sheets of RovoPromptSpec); anything else is a problem, never written.
    Cov "PRM40"
    Dim res As Object, items As New Collection, problems As New Collection, recs As Collection, rec As Variant
    Dim cSheet As Long, cRow As Long, cCol As Long, cVal As Long, allowed As Object, line As Long, sh As String, r As Long
    Dim v As String, col As String, spec As Variant, protected As Object
    Set res = CreateObject("Scripting.Dictionary")
    Set res("items") = items: Set res("problems") = problems: res("unknown") = 0: res("empty") = 0
    Set recs = ParseCsv(text)
    If recs.Count = 0 Then Cov "PRM41": problems.Add "The file has no line.": Set RovoCsvRead = res: Exit Function
    cSheet = CsvColumn(recs(1), "Sheet"): cRow = CsvColumn(recs(1), "Row"): cCol = CsvColumn(recs(1), "Column"): cVal = CsvColumn(recs(1), "Value")
    If cSheet < 0 Or cRow < 0 Or cCol < 0 Or cVal < 0 Then
        Cov "PRM42"
        problems.Add "The first line must name the columns Sheet, Row, Column and Value (it holds: " & Join(recs(1), ", ") & ")."
        Set RovoCsvRead = res: Exit Function
    End If
    Set allowed = CreateObject("Scripting.Dictionary"): Set protected = CreateObject("Scripting.Dictionary")
    Set res("protected") = protected
    allowed("config") = True
    For Each spec In ReadTable(wb.Worksheets(SheetName(SHEET_ROVO_SPEC)))
        Select Case LCase$(CStr(spec("Part") & ""))                 ' a sheet of the prompt, or one the CSV may fill though the prompt does not ask for it
            Case "sheet", "load": allowed(LCase$(LogicalSheet(CStr(spec("Name") & "")))) = True     ' the spec names the real sheets
        End Select
        If LCase$(CStr(spec("Part") & "")) = "keep" Then protected(LCase$(CStr(spec("Name") & ""))) = True
    Next spec
    For Each rec In recs
        line = line + 1
        If line > 1 Then
            sh = LogicalSheet(TrimWhite(CsvAt(rec, cSheet))): col = TrimWhite(CsvAt(rec, cCol)): v = TrimWhite(CsvAt(rec, cVal))
            If Len(sh) = 0 And Len(col) = 0 And Len(v) = 0 Then
                ' a blank line of the table
            ElseIf Not allowed.Exists(LCase$(sh)) Then
                Cov "PRM43": problems.Add "Line " & line & ": the sheet '" & sh & "' is not one that this import fills."
            ElseIf Len(col) = 0 Then
                problems.Add "Line " & line & ": the column is empty."
            Else
                ImportItem items, problems, res, protected, line, sh, TrimWhite(CsvAt(rec, cRow)), col, v
            End If
        End If
    Next rec
    Set RovoCsvRead = res
End Function

Private Sub ImportItem(ByVal items As Collection, ByVal problems As Collection, ByVal res As Object, ByVal protected As Object, ByVal line As Long, _
                       ByVal sh As String, ByVal rowText As String, ByVal col As String, ByVal v As String)
    Dim isConfig As Boolean, r As Long, noValue As Boolean
    isConfig = (StrComp(sh, "Config", vbTextCompare) = 0)
    If isConfig Then
        If Not IsPlainAscii(col) Then
            Cov "PRM51"
            problems.Add "Line " & line & ": a Config parameter name must be plain ASCII letters and digits."
            Exit Sub
        End If
        If protected.Exists(LCase$(col)) Then
            Cov "PRM50"
            problems.Add "Line " & line & ": " & col & " says where the credentials go: type it by hand, it is never loaded from a CSV."
            Exit Sub
        End If
    End If
    If Len(v) = 0 Then
        Cov "PRM44": res("empty") = res("empty") + 1: noValue = True       ' an empty cell, as asked when there is no value
    ElseIf StrComp(v, "UNKNOWN", vbTextCompare) = 0 Or StrComp(Left$(v, 12), "TO BE TYPED ", vbTextCompare) = 0 Then
        Cov "PRM53": res("unknown") = res("unknown") + 1: noValue = True: v = ""
    End If
    If isConfig Then
        If Not noValue Then items.Add Array("Config", 0, col, v)
        Exit Sub
    End If
    If Not IsNumeric(rowText) Or Len(rowText) = 0 Then
        Cov "PRM45": problems.Add "Line " & line & ": Row must be a number for the sheet " & sh & " (got '" & rowText & "')."
        Exit Sub
    End If
    r = CLng(Val(rowText))
    If r < 1 Then problems.Add "Line " & line & ": Row must be 1 or more.": Exit Sub
    items.Add Array(sh, r, col, v)
End Sub

Private Function IsPlainAscii(ByVal text As String) As Boolean
    Dim i As Long, c As Long
    For i = 1 To Len(text)
        c = AscW(Mid$(text, i, 1))
        If c < 32 Or c > 126 Then Exit Function
    Next i
    IsPlainAscii = True
End Function

Private Function CsvColumn(ByVal rec As Variant, ByVal name As String) As Long
    Dim i As Long
    CsvColumn = -1
    For i = LBound(rec) To UBound(rec)
        If StrComp(TrimWhite(CStr(rec(i))), name, vbTextCompare) = 0 Then CsvColumn = i: Exit Function
    Next i
End Function

Private Function CsvAt(ByVal rec As Variant, ByVal i As Long) As String
    If i <= UBound(rec) Then CsvAt = CStr(rec(i))
End Function

Public Function HeaderColumn(ByVal ws As Object, ByVal name As String) As Long
    ' The column of a header of row 1 (0 when missing); the cells after the first empty header are not read.
    Dim c As Long
    For c = 1 To ws.Columns.Count
        If Len(CStr(ws.Cells(1, c).Value & "")) = 0 Then Exit Function
        If StrComp(CStr(ws.Cells(1, c).Value), name, vbTextCompare) = 0 Then HeaderColumn = c: Exit Function
    Next c
End Function

Public Function RovoCsvApply(ByVal wb As Object, ByVal res As Object) As String
    ' Writes the items into the parameter sheets and returns the summary for the status cell. Config: the Value of the parameter
    ' named in Column (the other cells are not touched; a parameter that is not in Config is a problem). Any other sheet: all its
    ' rows below the header are replaced by the rows of the CSV (the column Example is kept), cell by cell: a whole number is a
    ' number, "-" is empty, the rest is text (a date stays ISO text, as in the sheets of the repository).
    Cov "PRM46"
    Dim item As Variant, ws As Object, counts As Object, cleared As Object, names As String, k As Variant, c As Long, r As Long
    Dim cValue As Long, cParam As Long, last As Long, problems As Collection, summary As String, nProblems As Long, first As String
    Dim loaded As String
    Set counts = CreateObject("Scripting.Dictionary"): Set cleared = CreateObject("Scripting.Dictionary")
    Set problems = res("problems")
    For Each item In res("items")
        If Not SheetExists(wb, CStr(item(0))) Then
            problems.Add "The sheet " & item(0) & " is not in the workbook."
        Else
            Set ws = wb.Worksheets(SheetName(CStr(item(0))))
            If StrComp(CStr(item(0)), "Config", vbTextCompare) = 0 Then
                cParam = HeaderColumn(ws, "Parameter"): cValue = HeaderColumn(ws, "Value")
                r = FindRow(ws, cParam, CStr(item(2)))
                If r = 0 Or cValue = 0 Then
                    problems.Add SheetName("Config") & " has no parameter " & item(2) & "."
                ElseIf res("protected").Exists(LCase$(TrimWhite(CStr(ws.Cells(r, cParam).Value & "")))) Then
                    Cov "PRM52"
                    problems.Add SheetName("Config") & " parameter " & TrimWhite(CStr(ws.Cells(r, cParam).Value & "")) & " says where the credentials go: type it by hand."
                Else
                    ws.Cells(r, cValue).Value = SafeCell(CStr(item(3)))
                    counts("Config") = Val(counts("Config") & "") + 1
                    If counts("Config") <= 8 Then loaded = loaded & IIf(Len(loaded) > 0, ", ", "") & item(2)
                End If
            Else
                If Not cleared.Exists(ws.Name) Then ClearTableBody ws: cleared(ws.Name) = 0
                c = HeaderColumn(ws, CStr(item(2)))
                If c = 0 Then
                    Cov "PRM47": problems.Add "The sheet " & ws.Name & " has no column " & item(2) & "."
                Else
                    ws.Cells(CLng(item(1)) + 1, c).Value = CsvCell(CStr(item(3)))
                    If CLng(item(1)) > cleared(ws.Name) Then cleared(ws.Name) = CLng(item(1))
                End If
            End If
        End If
    Next item
    For Each k In cleared.Keys
        names = names & " | " & k & ": " & cleared(k) & " row(s)"
    Next k
    ' "Title | Name: value | ...": the Dashboard shows the title in its status cell and one row per field under it.
    summary = "Settings CSV loaded"
    If counts.Exists("Config") Then summary = summary & " | " & SheetName("Config") & ": " & counts("Config") & " value(s): " & loaded & IIf(counts("Config") > 8, ", ...", "")
    summary = summary & names
    If counts.Count = 0 And Len(names) = 0 Then summary = summary & " | Result: nothing to load"
    summary = summary & " | Empty values: " & res("empty") & " (not loaded)" & " | UNKNOWN values: " & res("unknown") & " (not loaded)"
    nProblems = problems.Count
    If nProblems > 0 Then
        Cov "PRM48"
        first = CStr(problems(1))
        summary = summary & " | Problems: " & nProblems & ", first: " & first
    End If
    RovoCsvApply = summary
End Function

Public Function FindRow(ByVal ws As Object, ByVal col As Long, ByVal name As String) As Long
    Dim r As Long, last As Long
    If col = 0 Then Exit Function
    last = ws.Cells(ws.Rows.Count, col).End(-4162).Row
    For r = 2 To last
        If StrComp(TrimWhite(CStr(ws.Cells(r, col).Value & "")), name, vbTextCompare) = 0 Then FindRow = r: Exit Function
    Next r
End Function

Private Sub ClearTableBody(ByVal ws As Object)
    ' Every row below the header, in the columns of the table (the column Example, a note for the user, stays).
    Dim c As Long, last As Long, head As String
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If last < 2 Then Exit Sub
    For c = 1 To ws.Columns.Count
        head = CStr(ws.Cells(1, c).Value & "")
        If Len(head) = 0 Then Exit For
        If head <> "Example" Then ws.Range(ws.Cells(2, c), ws.Cells(last, c)).ClearContents
    Next c
End Sub

Private Function CsvCell(ByVal v As String) As Variant
    ' "-" is none (as in params\*.md), digits only is a whole number, the rest is text.
    Dim i As Long, digits As Boolean
    If v = "-" Or Len(v) = 0 Then CsvCell = Empty: Exit Function
    digits = (Len(v) < 10)
    For i = 1 To Len(v)
        If Mid$(v, i, 1) < "0" Or Mid$(v, i, 1) > "9" Then digits = False: Exit For
    Next i
    If digits Then CsvCell = CLng(v) Else CsvCell = SafeCell(v)
End Function

Public Function RovoCsvText(ByVal path As String) As String
    ' The text of the file, UTF-8 (a BOM is dropped).
    Cov "PRM49"
    Dim f As Integer, bytes() As Byte, size As Long
    If Len(Dir$(path)) = 0 Then Err.Raise vbObjectError + 631, "RovoCsvText", "The file " & path & " does not exist."
    f = FreeFile
    Open path For Binary Access Read As #f
    size = LOF(f)
    If size > 0 Then
        ReDim bytes(0 To size - 1)
        Get #f, , bytes
    End If
    Close #f
    If size > 0 Then RovoCsvText = Utf8Decode(bytes)
End Function

Public Function ImportRovoCsv(ByVal wb As Object, ByVal path As String) As String
    ' The answer of Rovo (CSV) into the parameter sheets; returns the summary. The file may be in any folder (the button asks
    ' for it in a dialog). Nothing is sent anywhere: the file is read here.
    ImportRovoCsv = RovoCsvApply(wb, RovoCsvRead(wb, RovoCsvText(path)))
End Function

Public Function IsYes(ByVal text As String) As Boolean
    ' The sheet is read in a French Excel: Yes / Oui / True / Vrai / 1 mean yes.
    Cov "PRM17"
    Select Case LCase$(TrimWhite(text))
        Case "yes", "oui", "true", "vrai", "1": IsYes = True
    End Select
End Function

' ================= The status rules that come from Jira (flow/status_rules.py: the same rules, the same cases in T_StatusRules) =================

Public Function JiraCategoryName(ByVal key As String) As String
    ' Backlog, Active or Done for a status category of Jira (new, indeterminate, done; or the enum TODO, IN_PROGRESS, DONE), "" when unknown.
    Select Case LCase$(TrimWhite(key))
        Case "new", "todo", "to do": Cov "PRM56": JiraCategoryName = "Backlog"
        Case "indeterminate", "in_progress", "in progress": Cov "PRM57": JiraCategoryName = "Active"
        Case "done": Cov "PRM58": JiraCategoryName = "Done"
    End Select
End Function

Public Function FirstInProgress(ByVal initial As String, ByVal moves As Collection, ByVal categories As Object) As String
    ' The start of a workflow: the first status in progress (Jira category indeterminate) that a breadth-first search reaches from the status of creation
    ' along moves (Array(from, to)): the nearest wins, the first listed breaks a tie; "" when nothing is reached. categories: status -> Jira category.
    Cov "PRM59"
    Dim graph As Object, seen As Object, queue As New Collection, mv As Variant, status As String, target As Variant, head As Long
    If Len(initial) = 0 Then Exit Function
    Set graph = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary")
    For Each mv In moves
        If Not graph.Exists(CStr(mv(0))) Then Set graph(CStr(mv(0))) = New Collection
        graph(CStr(mv(0))).Add CStr(mv(1))
    Next mv
    seen(initial) = True: queue.Add initial: head = 1
    Do While head <= queue.Count
        status = CStr(queue(head)): head = head + 1
        If categories.Exists(status) Then
            If JiraCategoryName(CStr(categories(status))) = "Active" Then Cov "PRM60": FirstInProgress = status: Exit Function
        End If
        If graph.Exists(status) Then
            For Each target In graph(status)
                If Not seen.Exists(CStr(target)) Then seen(CStr(target)) = True: queue.Add CStr(target)
            Next target
        End If
    Loop
End Function

Public Function FlowDepths(ByVal initial As String, ByVal moves As Collection) As Object
    ' status -> the number of moves from the creation status (breadth-first search over moves, Array(from, to)): 0 for the creation status. Empty when there is
    ' no creation status. Twin of flow/status_rules.flow_depths.
    Cov "PRM68"
    Dim out As Object, graph As Object, queue As New Collection, mv As Variant, status As String, target As Variant, head As Long
    Set out = CreateObject("Scripting.Dictionary")
    Set FlowDepths = out
    If Len(initial) = 0 Then Exit Function
    Set graph = CreateObject("Scripting.Dictionary")
    For Each mv In moves
        If Not graph.Exists(CStr(mv(0))) Then Set graph(CStr(mv(0))) = New Collection
        graph(CStr(mv(0))).Add CStr(mv(1))
    Next mv
    out(initial) = 0: queue.Add initial: head = 1
    Do While head <= queue.Count
        status = CStr(queue(head)): head = head + 1
        If graph.Exists(status) Then
            For Each target In graph(status)
                If Not out.Exists(CStr(target)) Then out(CStr(target)) = out(status) + 1: queue.Add CStr(target)
            Next target
        End If
    Loop
End Function

Private Function CellNumber(ByVal cell As Variant, ByRef value As Long) As Boolean
    ' True (and the whole number in value) when a cell of the sheet Workflows is filled.
    If Len(TrimWhite(CStr(cell & ""))) = 0 Then Exit Function
    value = CLng(Val(CStr(cell)))
    CellNumber = True
End Function

Public Function EffectiveRules(ByVal kinds As Variant, ByVal catalog As Object, ByVal currentOf As Object, ByVal starts As Object, _
                               ByVal frees As Object, ByVal exceptions As Collection) As Collection
    ' One rule per workflow kind and status: Array(kind, status, category, isStart, isEnd, isFree, source), the kinds in the order given then the names in
    ' binary order. catalog: status -> Jira category, currentOf: the category Jira gives the tickets' current statuses, starts / frees: dictionaries
    ' of the statuses the Jira workflows start with / reach freely. exceptions: Array(workflow, status, category, isStart, isEnd, isFree) of the sheet
    ' Workflows, an empty cell is Empty or "": it takes Jira's value; a row without workflow touches both, a row of the workflow comes after it.
    Cov "PRM61"
    Dim names As Object, nm As Variant, ex As Variant, kind As Variant, sorted() As String, n As Long, i As Long, j As Long, tmp As String
    Dim out As New Collection, cat As String, isStart As Long, isFree As Long, isEnd As Long, hasEnd As Boolean, applied As Boolean, pass As Long, v As Long
    Set names = CreateObject("Scripting.Dictionary")
    For Each nm In catalog.Keys: names(CStr(nm)) = True: Next nm
    For Each nm In currentOf.Keys: names(CStr(nm)) = True: Next nm
    For Each ex In exceptions: names(CStr(ex(1))) = True: Next ex
    Set EffectiveRules = out
    n = names.Count
    If n = 0 Then Exit Function
    ReDim sorted(1 To n)
    For Each nm In names.Keys
        i = i + 1: sorted(i) = CStr(nm)
    Next nm
    For i = 2 To n                                                   ' insertion sort, binary compare: the order of Python's sorted on the names
        tmp = sorted(i): j = i - 1
        Do While j >= 1
            If StrComp(sorted(j), tmp, vbBinaryCompare) <= 0 Then Exit Do
            sorted(j + 1) = sorted(j): j = j - 1
        Loop
        sorted(j + 1) = tmp
    Next i
    For Each kind In kinds
        For i = 1 To n
            nm = sorted(i): cat = ""
            If catalog.Exists(nm) Then cat = JiraCategoryName(CStr(catalog(nm)))
            If Len(cat) = 0 And currentOf.Exists(nm) Then Cov "PRM62": cat = JiraCategoryName(CStr(currentOf(nm)))
            If Len(cat) = 0 Then Cov "PRM63": cat = "Unclassified"
            isStart = IIf(starts.Exists(nm), 1, 0): isFree = IIf(frees.Exists(nm), 1, 0): hasEnd = False: applied = False
            For pass = 0 To 1                                        ' the rows of both workflows, then those of this workflow (the more precise wins)
                For Each ex In exceptions
                    If CStr(ex(1)) = CStr(nm) And IIf(pass = 0, Len(CStr(ex(0) & "")) = 0, CStr(ex(0) & "") = CStr(kind)) Then
                        Cov "PRM64"
                        applied = True
                        If Len(CStr(ex(2) & "")) > 0 Then cat = CStr(ex(2))
                        If CellNumber(ex(3), v) Then isStart = v
                        If CellNumber(ex(4), v) Then isEnd = v: hasEnd = True
                        If CellNumber(ex(5), v) Then isFree = v
                    End If
                Next ex
            Next pass
            If Not hasEnd Then isEnd = IIf(cat = "Done", 1, 0)
            out.Add Array(CStr(kind), CStr(nm), cat, isStart, isEnd, isFree, IIf(applied, "user", "Jira"))
        Next i
    Next kind
End Function

Public Function RuleCategories(ByVal rules As Collection, ByVal kind As String) As Object
    ' status -> category of the rules of one workflow kind: what BuildStatusIntervals reads.
    Dim out As Object, r As Variant
    Set out = CreateObject("Scripting.Dictionary")
    For Each r In rules
        If CStr(r(0)) = kind Then out(CStr(r(1))) = CStr(r(2))
    Next r
    Set RuleCategories = out
End Function

Public Function HasWait(ByVal rules As Collection) As Boolean
    ' True once a status is declared Wait: without one Flow Efficiency would be 100 % by construction, so it is not computable.
    Dim r As Variant
    For Each r In rules
        If CStr(r(2)) = "Wait" Then Cov "PRM65": HasWait = True: Exit Function
    Next r
End Function

' What the rules hold, and the warnings when a metric cannot be computed (flow/status_rules.py: rule_counts, counts_line, rule_warnings).

Private Function KindCount(ByVal rules As Collection, ByVal kind As String, ByVal what As String) As Long
    ' The rules of one kind that start the flow (what = "start"), end it ("end"), or have the category what.
    Dim r As Variant, n As Long
    For Each r In rules
        If CStr(r(0)) = kind Then
            Select Case what
                Case "start": If Val(CStr(r(3) & "")) = 1 Then n = n + 1
                Case "end": If Val(CStr(r(4) & "")) = 1 Then n = n + 1
                Case Else: If CStr(r(2)) = what Then n = n + 1
            End Select
        End If
    Next r
    KindCount = n
End Function

Public Function RuleCountsLine(ByVal rules As Collection, ByVal kinds As Variant) As String
    ' The INFO line of the log: "status rules | Feature: start 1, active 2, wait 0, end 2 | Team: ...".
    Cov "PRM69"
    Dim kind As Variant, out As String
    For Each kind In kinds
        out = out & " | " & CStr(kind) & ": start " & KindCount(rules, CStr(kind), "start") & ", active " & KindCount(rules, CStr(kind), "Active") & _
              ", wait " & KindCount(rules, CStr(kind), "Wait") & ", end " & KindCount(rules, CStr(kind), "end")
    Next kind
    RuleCountsLine = "status rules" & out
End Function

Public Function RuleWarnings(ByVal rules As Collection, ByVal kinds As Variant, ByVal withTickets As Object) As Collection
    ' The WARN lines: a kind that has tickets (withTickets: kind -> True) and no start status (no StartUtc, so no Flow Time, Flow Efficiency or Flow
    ' Load), then no Wait status in any kind (Flow Efficiency reads the whole sheet StatusRules). None when no kind has a ticket.
    Cov "PRM70"
    Dim out As New Collection, kind As Variant
    For Each kind In kinds
        If withTickets.Exists(CStr(kind)) Then
            If KindCount(rules, CStr(kind), "start") = 0 Then Cov "PRM71": out.Add "J2D-FLOW-010 | no start status in the " & CStr(kind) & NO_START_TEXT
        End If
    Next kind
    If withTickets.Count > 0 Then
        If Not HasWait(rules) Then Cov "PRM72": out.Add "J2D-FLOW-011 | " & NO_WAIT_TEXT
    End If
    Set RuleWarnings = out
End Function

Public Function WorkflowExceptions(ByVal wsWorkflows As Object) As Collection
    ' The rows of the sheet Workflows as the exceptions of EffectiveRules: Array(workflow, status, category, isStart, isEnd, isFree). Unlike ReadTable it also
    ' reads a row whose Workflow is empty (it touches both workflows); a row without a status is left out.
    Cov "PRM66"
    Dim out As New Collection, cW As Long, cS As Long, cC As Long, cB As Long, cE As Long, cF As Long, last As Long, r As Long, st As String
    Set WorkflowExceptions = out
    cS = HeaderColumn(wsWorkflows, "Status")
    If cS = 0 Then Exit Function
    cW = HeaderColumn(wsWorkflows, "Workflow"): cC = HeaderColumn(wsWorkflows, "Category")
    cB = HeaderColumn(wsWorkflows, "IsStart"): cE = HeaderColumn(wsWorkflows, "IsEnd"): cF = HeaderColumn(wsWorkflows, "IsFree")
    last = wsWorkflows.Cells(wsWorkflows.Rows.Count, cS).End(-4162).Row
    For r = 2 To last
        st = TrimWhite(CStr(wsWorkflows.Cells(r, cS).Value & ""))
        If Len(st) > 0 Then out.Add Array(CellText(wsWorkflows, r, cW), st, CellText(wsWorkflows, r, cC), CellText(wsWorkflows, r, cB), _
                                         CellText(wsWorkflows, r, cE), CellText(wsWorkflows, r, cF))
    Next r
End Function

Private Function CellText(ByVal ws As Object, ByVal r As Long, ByVal c As Long) As String
    If c = 0 Then Exit Function
    CellText = TrimWhite(CStr(ws.Cells(r, c).Value & ""))
End Function

Public Function FlaggedStatuses(ByVal ws As Object, ByVal flagHeader As String) As Object
    ' status -> True for the rows of a sheet (Status column) whose flag column holds 1: the starts and the free statuses of the sheet JiraWorkflows.
    Dim out As Object, row As Variant
    Set out = CreateObject("Scripting.Dictionary")
    For Each row In ReadTable(ws)
        If Val(CStr(row(flagHeader) & "")) = 1 Then Cov "PRM67": out(CStr(row("Status") & "")) = True
    Next row
    Set FlaggedStatuses = out
End Function
