Option Explicit
' What the person should look at after an extraction: fields " | Name: value" for the status cell, each one only when there is
' something to see (a clean run adds nothing). They say WHY a figure or a chart is empty, with numbers and the names of the
' statuses (the person's own configuration), never a value of a ticket:
'   Done outside the iterations   tickets whose DoneUtc falls in no iteration of the sheet Iterations (no metric counts them)
'   Without a team                tickets whose Team is empty (the Team comes from the sheet Projects)
'   Statuses not in Workflows     statuses of the history that the sheet Workflows lacks (category Unclassified: the Flow
'                                 Efficiency is not computable and the Done of the ticket is not known)
'   Cells in error                Table.Column x n for the first columns of the flow tables that show an Excel error
' Reads the tables of the workbook; never calls Jira, never raises (a status must not hide a finished extraction).

Private Const MAX_NAMES As Long = 5

Public Function HealthFields(ByVal wb As Object) As String
    Cov "HLT01"
    On Error GoTo Failed
    HealthFields = OutsideIterations(wb) & WithoutTeam(wb) & UnclassifiedStatuses(wb) & CellsInError(wb)
    Exit Function
Failed:
    Cov "HLT02"
    HealthFields = ""
End Function

Private Function TableOf(ByVal wb As Object, ByVal name As String) As Object
    ' The table `name` of the sheet of the same logical name; Nothing when the workbook has none.
    If Not SheetExists(wb, name) Then Exit Function
    On Error Resume Next
    Set TableOf = wb.Worksheets(SheetName(name)).ListObjects(name)
    On Error GoTo 0
End Function

Private Function ColumnValues(ByVal lo As Object, ByVal header As String) As Variant
    ' The cells of a column of a table as a 2-D array (rows, 1); Empty when the table has no row or no such column.
    Dim col As Object, one(1 To 1, 1 To 1) As Variant
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    On Error Resume Next
    Set col = lo.ListColumns(header)
    On Error GoTo 0
    If col Is Nothing Then Exit Function
    If col.DataBodyRange.Rows.Count = 1 Then
        one(1, 1) = col.DataBodyRange.Value                              ' one row: Value is a single value, not an array
        ColumnValues = one
    Else
        ColumnValues = col.DataBodyRange.Value
    End If
End Function

Private Function IsDoneDate(ByVal v As Variant) As Boolean
    ' DoneUtc is a date serial when the ticket is done, "" (or an error) when it is not.
    If IsError(v) Then Exit Function
    If Len(CStr(v & "")) = 0 Then Exit Function
    IsDoneDate = IsNumeric(v)
End Function

Private Function InAnIteration(ByVal serial As Double, ByVal starts As Collection, ByVal ends As Collection) As Boolean
    Dim k As Long
    For k = 1 To starts.Count
        If serial >= starts(k) And serial < ends(k) Then InAnIteration = True: Exit Function
    Next k
End Function

Private Function OutsideIterations(ByVal wb As Object) As String
    Dim done As Variant, starts As New Collection, ends As New Collection, row As Variant, s As Variant, e As Variant
    Dim r As Long, nDone As Long, nOut As Long
    If Not SheetExists(wb, "Iterations") Then Exit Function
    done = ColumnValues(TableOf(wb, "ItemFlow"), "DoneUtc")
    If IsEmpty(done) Then Exit Function
    For Each row In ReadTable(wb.Worksheets(SheetName("Iterations")))
        s = SerialOf(row("Start")): e = SerialOf(row("End"))
        If Not IsEmpty(s) And Not IsEmpty(e) Then starts.Add s: ends.Add e
    Next row
    For r = 1 To UBound(done, 1)
        If IsDoneDate(done(r, 1)) Then
            nDone = nDone + 1
            If Not InAnIteration(CDbl(done(r, 1)), starts, ends) Then nOut = nOut + 1
        End If
    Next r
    If nOut > 0 Then
        Cov "HLT03"
        OutsideIterations = " | Done outside the iterations: " & nOut & " of " & nDone & " done tickets (no metric counts them): " & _
                            "add iterations to " & SheetName("Iterations")
    End If
End Function

Private Function WithoutTeam(ByVal wb As Object) As String
    Dim team As Variant, r As Long, n As Long
    team = ColumnValues(TableOf(wb, "ItemFlow"), "Team")
    If IsEmpty(team) Then Exit Function
    For r = 1 To UBound(team, 1)
        If IsError(team(r, 1)) Then
            n = n + 1
        ElseIf Len(TrimWhite(CStr(team(r, 1) & ""))) = 0 Then
            n = n + 1
        End If
    Next r
    If n > 0 Then
        Cov "HLT04"
        WithoutTeam = " | Without a team: " & n & " tickets (a ticket takes its team from " & SheetName("Projects") & ")"
    End If
End Function

Private Function UnclassifiedStatuses(ByVal wb As Object) As String
    ' The distinct statuses of StatusIntervals whose Category is Unclassified.
    Dim status As Variant, cat As Variant, seen As Object, r As Long, names As String, n As Long, name As String
    status = ColumnValues(TableOf(wb, "StatusIntervals"), "Status")
    cat = ColumnValues(TableOf(wb, "StatusIntervals"), "Category")
    If IsEmpty(status) Or IsEmpty(cat) Then Exit Function
    Set seen = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(status, 1)
        If Not IsError(status(r, 1)) And Not IsError(cat(r, 1)) Then
            name = CStr(status(r, 1) & "")
            If CStr(cat(r, 1) & "") = "Unclassified" And Not seen.Exists(name) Then
                seen(name) = True
                n = n + 1
                If n <= MAX_NAMES Then names = names & IIf(Len(names) > 0, ", ", "") & name
            End If
        End If
    Next r
    If n > 0 Then
        Cov "HLT05"
        UnclassifiedStatuses = " | Statuses not in Workflows: " & n & " (" & names & IIf(n > MAX_NAMES, ", ...", "") & "): add them to " & _
                               SheetName("Workflows") & " with a Category"
    End If
End Function

Private Function ErrorCount(ByVal ws As Object, ByVal col As Object) As Double
    ' How many cells of the column show an Excel error.
    Dim n As Variant
    n = Application.Evaluate("SUMPRODUCT(--ISERROR('" & ws.Name & "'!" & col.DataBodyRange.Address & "))")
    If Not IsError(n) Then ErrorCount = CDbl(n)
End Function

Private Function CellsInError(ByVal wb As Object) As String
    ' Table.Column x n, for the first columns of the flow tables that hold an Excel error (the charts' own #N/A are gaps on purpose
    ' and are not read here). A column with an error is a formula that fails on the data.
    Dim name As Variant, lo As Object, col As Object, n As Double, found As String, columns As Long, shown As Long
    For Each name In Array("StatusIntervals", "ItemFlow", "DailyLoad", "Metrics", "PIMetrics", "DataChecks", "FillRates", "IssueLinks")
        Set lo = TableOf(wb, CStr(name))
        If Not lo Is Nothing Then
            If Not lo.DataBodyRange Is Nothing Then
                For Each col In lo.ListColumns
                    n = ErrorCount(wb.Worksheets(SheetName(CStr(name))), col)
                    If n > 0 Then
                        columns = columns + 1
                        If shown < MAX_NAMES Then
                            shown = shown + 1
                            found = found & IIf(Len(found) > 0, "; ", "") & CStr(name) & "." & col.Name & " x " & n
                        End If
                    End If
                Next col
            End If
        End If
    Next name
    If columns > 0 Then
        Cov "HLT06"
        CellsInError = " | Cells in error: " & columns & " column(s), first " & found & IIf(columns > shown, "; ...", "")
    End If
End Function
