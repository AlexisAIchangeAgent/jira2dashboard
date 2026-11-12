Attribute VB_Name = "modParams"
Option Explicit
' Parameters live in the workbook, never in code. The sheets Workflows, Transitions and FieldMapping are written by
' the Python generator from params\*.md; this module only reads them. Each reader takes the sheet as an argument
' (testable on a scratch sheet). Every call re-reads the sheet: keep the result (for example the Dictionary of
' StatusCategories) when you need it in a loop.

Public Const SHEET_WORKFLOWS As String = "Workflows"
Public Const SHEET_TRANSITIONS As String = "Transitions"
Public Const SHEET_ROVO_SPEC As String = "RovoPromptSpec"
Public Const SHEET_ROVO As String = "RovoPrompt"

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
    Dim s As String, d As Date, sec As Long
    ParseMoment = Empty
    s = TrimWhite(text)
    If Len(s) < 10 Then Exit Function
    If Mid$(s, 5, 1) <> "-" Or Mid$(s, 8, 1) <> "-" Then Exit Function
    On Error GoTo Invalid
    d = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2)))
    If Month(d) <> CInt(Mid$(s, 6, 2)) Or Day(d) <> CInt(Mid$(s, 9, 2)) Then GoTo Invalid     ' 2026-13-45 does not exist
    If Len(s) >= 16 Then
        If Len(s) >= 19 Then sec = CInt(Mid$(s, 18, 2))
        d = d + TimeSerial(CInt(Mid$(s, 12, 2)), CInt(Mid$(s, 15, 2)), sec)
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
        If CDbl(v) < 36526 Then                      ' before 2000-01-01: a plain number such as 5 was typed, not a date
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

Public Function RovoPromptLines(ByVal wb As Object) As Collection
    ' The prompt for Rovo, one line per item: the rows of the sheet RovoPromptSpec (Part | Name | Text) in order. A text row
    ' is a line as it is, "config" becomes one line per parameter of the sheet Config, "sheet" one line for a parameter sheet.
    ' The Value column of Config is never read: it may hold the address, the e-mail or the token.
    Cov "PRM30"
    Dim out As New Collection, spec As Variant
    For Each spec In ReadTable(wb.Worksheets(SHEET_ROVO_SPEC))
        Select Case LCase$(CStr(spec("Part") & ""))
            Case "keep"                                                  ' a connection parameter: not part of the prompt text
            Case "config": Cov "PRM31": AddConfigLines wb.Worksheets("Config"), out
            Case "sheet": Cov "PRM32": out.Add SheetLine(wb, CStr(spec("Name") & ""), CStr(spec("Text") & ""))
            Case Else: Cov "PRM33": out.Add CStr(spec("Text") & "")
        End Select
    Next spec
    Set RovoPromptLines = out
End Function

Private Sub AddConfigLines(ByVal wsConfig As Object, ByVal out As Collection)
    ' "- Name (mandatory): meaning. Example: form of the value", from the columns Parameter, Mandatory, Description, Example.
    Dim row As Variant, text As String, example As String
    For Each row In ReadTable(wsConfig)
        text = "- " & CStr(row("Parameter") & "") & IIf(IsYes(CStr(row("Mandatory") & "")), " (mandatory): ", " (optional): ") & _
               CStr(row("Description") & "")
        example = CStr(row("Example") & "")
        If Len(example) > 0 Then text = text & " Example: " & example
        out.Add text
    Next row
End Sub

Private Function SheetLine(ByVal wb As Object, ByVal name As String, ByVal rule As String) As String
    ' "SHEET Projects - columns: Project | Role | Team. rule", the columns read from the header row of the sheet.
    Dim ws As Object, c As Long, cols As String, head As String
    If Not SheetExists(wb, name) Then SheetLine = "SHEET " & name & " - this sheet is not in the workbook. " & rule: Exit Function
    Set ws = wb.Worksheets(name)
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
        Set ws = wb.Worksheets(SHEET_ROVO)
    Else
        Cov "PRM35"
        Set ws = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SHEET_ROVO
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
    For Each spec In ReadTable(wb.Worksheets(SHEET_ROVO_SPEC))
        If LCase$(CStr(spec("Part") & "")) = "sheet" Then allowed(LCase$(CStr(spec("Name") & ""))) = True
        If LCase$(CStr(spec("Part") & "")) = "keep" Then protected(LCase$(CStr(spec("Name") & ""))) = True
    Next spec
    For Each rec In recs
        line = line + 1
        If line > 1 Then
            sh = TrimWhite(CsvAt(rec, cSheet)): col = TrimWhite(CsvAt(rec, cCol)): v = TrimWhite(CsvAt(rec, cVal))
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

Private Function HeaderColumn(ByVal ws As Object, ByVal name As String) As Long
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
            Set ws = wb.Worksheets(CStr(item(0)))
            If StrComp(CStr(item(0)), "Config", vbTextCompare) = 0 Then
                cParam = HeaderColumn(ws, "Parameter"): cValue = HeaderColumn(ws, "Value")
                r = FindRow(ws, cParam, CStr(item(2)))
                If r = 0 Or cValue = 0 Then
                    problems.Add "Config has no parameter " & item(2) & "."
                ElseIf res("protected").Exists(LCase$(TrimWhite(CStr(ws.Cells(r, cParam).Value & "")))) Then
                    Cov "PRM52"
                    problems.Add "Config parameter " & TrimWhite(CStr(ws.Cells(r, cParam).Value & "")) & " says where the credentials go: type it by hand."
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
    If counts.Exists("Config") Then summary = summary & " | Config: " & counts("Config") & " value(s): " & loaded & IIf(counts("Config") > 8, ", ...", "")
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

Private Function FindRow(ByVal ws As Object, ByVal col As Long, ByVal name As String) As Long
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

