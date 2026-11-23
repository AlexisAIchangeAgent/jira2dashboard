Option Explicit
' What the person should look at after an extraction: fields " | Name: value" for the status cell, each one only when there is
' something to see (a clean run adds nothing). They say WHY a figure or a chart is empty, with numbers and the names of the
' statuses (the person's own configuration), never a value of a ticket. The Python twin is flow/health.py (same fields, same order, same
' words: tests/unit/test_health.py reads them here):
'   No start status               a workflow kind that has tickets and no IsStart status in StatusRules: no StartUtc, so no Flow Time,
'                                 Flow Efficiency or Flow Load (AL, 2026-11-22: "flow metrics are still not computable")
'   No Wait status                no status of StatusRules is Wait: Flow Efficiency is not computable
'   Not Features                  tickets of the Program that are not Features (team "ART (other types)" in ItemFlow): left out of the ART metrics;
'                                 and, when the ART has no Feature at all, that FeatureTypes matches none of them (a type named otherwise in Jira)
'   Done outside the iterations   tickets whose DoneUtc falls in no iteration of the sheet Iterations (no metric counts them)
'   Without a team                tickets whose Team is empty (the Team comes from the sheet Projects)
'   Statuses not in Workflows     statuses of the history that the sheet Workflows lacks (category Unclassified: the Flow
'                                 Efficiency is not computable and the Done of the ticket is not known)
' Started in the second of the backlog entry   tickets without a StartUtc that entered a start status in the same second as their last
'                                 backlog entry (the milliseconds are dropped and the start must come strictly after it)
'   AsOf before the data          status changes entered after AsOf: the open intervals and Flow Load stop at AsOf
'   VelocityCheck gap             rows of the table VelocityCheck whose Gap is not 0
'   Cells in error                Table.Column x n for the first columns of the flow tables that show an Excel error
'   Column types                  the first filled cell of a column written from Jira is not of the type of its label (TypeProblems)
' Reads the tables of the workbook; never calls Jira, never raises (a status must not hide a finished extraction).

Private Const MAX_NAMES As Long = 5
Private Const PROGRAM_TEAM As String = "ART"             ' the team of the Features of the Program (flow/model.ART)
Private Const OTHER_TEAM As String = "ART (other types)"  ' the team of the other tickets of the Program (flow/model.OTHER_TYPES)
Private Const FILL_LINE_MAX As Long = 270                ' a log line is cut at 300 characters: the fill counts go on a next line (raw_types.FILL_LINE_MAX)
Private Const BLOCK_ROWS As Long = 2000                  ' rows of a column read at once while the first filled cells are looked for
Private Const TYPE_SAMPLE As Long = 20                  ' the first filled cells of a column whose kinds the log counts (raw_types.SAMPLE)
Private Const FIRST_SERIAL As Double = 36526             ' 2000-01-01: a date before it is not one of Jira's (raw_types.DATE_FIRST)
Private Const LAST_SERIAL As Double = 73051              ' 2100-01-01 (raw_types.DATE_LAST)
Private Const HALF_SECOND As Double = 0.0000057870370370 ' a day is 1: half a second is the margin of a date at midnight

Public Function HealthFields(ByVal wb As Object) As String
    ' The fields in a step "health" of the log (after the step of TypeProblems); a failure keeps what was found and leaves its number in the log.
    Cov "HLT01"
    Dim found As String, begun As Boolean, num As Long, kinds As Object, starts As Object, waitFound As Boolean, ruled As Boolean
    On Error GoTo Failed
    found = TypeProblems(wb)
    LogBegin "health": begun = True
    Set kinds = KindsWithTickets(wb)
    Set starts = CreateObject("Scripting.Dictionary")
    ruled = RuleFacts(wb, starts, waitFound)
    found = found & NoStartStatus(kinds, starts, ruled) & NoWaitStatus(kinds, waitFound, ruled) & NotFeatures(wb) & OutsideIterations(wb) & _
            WithoutTeam(wb) & UnclassifiedStatuses(wb) & SameSecondStarts(wb) & AsOfBeforeData(wb) & VelocityCheckGap(wb) & ReconciliationFindings(wb) & _
            CellsInError(wb)
    LogEnd "health"
    HealthFields = found
    Exit Function
Failed:
    Cov "HLT02"
    num = Err.Number                                                    ' first: the calls below may clear Err
    LogLine "WARN", "health: the check failed | error " & num
    If begun Then LogEnd "health"
    HealthFields = found
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

Private Function SerialOfCell(ByVal v As Variant, ByRef serial As Double) As Boolean
    ' True (and the date serial) when a cell holds a date or a number; False for an error, a text or an empty cell.
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then Exit Function
    If VarType(v) = vbDate Or IsNumeric(v) Then serial = CDbl(v): SerialOfCell = True
End Function

Private Function KindsWithTickets(ByVal wb As Object) As Object
    ' The workflow kinds of the tickets of ItemFlow (column Workflow: Feature or Team); Nothing when no ticket was read (the one blank row of an empty
    ' table gives the kind Team in ItemFlow: the raw sheet Issues tells).
    Dim wf As Variant, r As Long, out As Object, ws As Object, headerRow As Long, firstCol As Long, nCols As Long, nRows As Long
    If Not TableShape(wb, "Issues", ws, headerRow, firstCol, nCols, nRows) Then Cov "HLT69": Exit Function
    wf = ColumnValues(TableOf(wb, "ItemFlow"), "Workflow")
    If IsEmpty(wf) Then Exit Function
    Set out = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(wf, 1)
        If Not IsError(wf(r, 1)) Then out(CStr(wf(r, 1) & "")) = True
    Next r
    Set KindsWithTickets = out
End Function

Private Function RuleFacts(ByVal wb As Object, ByVal starts As Object, ByRef waitFound As Boolean) As Boolean
    ' One read of the sheet StatusRules (a plain range: 2 rows per status of the instance): the kinds that have a start status (starts: kind -> True) and
    ' whether a status is Wait. False when the sheet or one of its columns is missing: then nothing is told about the rules.
    Dim ws As Object, cW As Long, cC As Long, cS As Long, last As Long, width As Long, data As Variant, r As Long
    If Not SheetExists(wb, "StatusRules") Then Exit Function
    Set ws = wb.Worksheets(SheetName("StatusRules"))
    cW = HeaderColumn(ws, "Workflow"): cC = HeaderColumn(ws, "Category"): cS = HeaderColumn(ws, "IsStart")
    If cW = 0 Or cC = 0 Or cS = 0 Then Exit Function
    RuleFacts = True
    last = ws.Cells(ws.Rows.Count, cW).End(-4162).Row                   ' xlUp
    If last < 2 Then Exit Function
    width = cW: If cC > width Then width = cC
    If cS > width Then width = cS
    data = ws.Range(ws.Cells(2, 1), ws.Cells(last, width)).Value        ' three columns at least: always a 2-D array
    For r = 1 To UBound(data, 1)
        If Not IsError(data(r, cW)) And Not IsError(data(r, cS)) Then
            If Val(CStr(data(r, cS) & "")) = 1 Then starts(CStr(data(r, cW) & "")) = True
        End If
        If Not IsError(data(r, cC)) Then
            If CStr(data(r, cC) & "") = "Wait" Then waitFound = True
        End If
    Next r
End Function

Private Function NoStartStatus(ByVal kinds As Object, ByVal starts As Object, ByVal ruled As Boolean) As String
    ' The workflow kinds that have tickets and no start status in the sheet StatusRules (flow/health.no_start_status).
    Dim kind As Variant, missing As String
    If kinds Is Nothing Or Not ruled Then Exit Function
    For Each kind In Array("Feature", "Team")
        If kinds.Exists(CStr(kind)) And Not starts.Exists(CStr(kind)) Then missing = missing & IIf(Len(missing) > 0, ", ", "") & CStr(kind)
    Next kind
    If Len(missing) > 0 Then
        Cov "HLT63"
        NoStartStatus = " | No start status: " & missing & " (no Flow Time, Flow Efficiency or Flow Load): set IsStart = 1 on the first in-progress status in " & _
                        SheetName("Workflows")
    End If
End Function

Private Function NoWaitStatus(ByVal kinds As Object, ByVal waitFound As Boolean, ByVal ruled As Boolean) As String
    ' No status of the sheet StatusRules is Wait while tickets were read: Flow Efficiency is not computable (flow/health.no_wait_status).
    If kinds Is Nothing Or Not ruled Or waitFound Then Exit Function
    If kinds.Count = 0 Then Exit Function
    Cov "HLT64"
    NoWaitStatus = " | No Wait status (no Flow Efficiency): set Category = Wait on the statuses where the work waits in " & SheetName("Workflows")
End Function

Private Function SameSecondStarts(ByVal wb As Object) As String
    ' Tickets without a StartUtc that entered a start status (StatusIntervals.IsStart) in the same second as their last backlog entry: the
    ' milliseconds are dropped and the start must come strictly after the backlog (flow/health.same_second_starts).
    Dim lo As Object, key As Variant, cat As Variant, entered As Variant, isStart As Variant, iKey As Variant, start As Variant
    Dim noStart As Object, last As Object, flagged As Object, r As Long, k As String, t As Double
    Set lo = TableOf(wb, "StatusIntervals")
    key = ColumnValues(lo, "Key"): cat = ColumnValues(lo, "Category"): entered = ColumnValues(lo, "EnteredUtc"): isStart = ColumnValues(lo, "IsStart")
    iKey = ColumnValues(TableOf(wb, "ItemFlow"), "Key"): start = ColumnValues(TableOf(wb, "ItemFlow"), "StartUtc")
    If IsEmpty(key) Or IsEmpty(cat) Or IsEmpty(entered) Or IsEmpty(isStart) Or IsEmpty(iKey) Or IsEmpty(start) Then Exit Function
    Set noStart = CreateObject("Scripting.Dictionary"): Set last = CreateObject("Scripting.Dictionary"): Set flagged = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(iKey, 1)
        If Not IsError(iKey(r, 1)) And Not IsError(start(r, 1)) Then
            If Len(CStr(start(r, 1) & "")) = 0 Then noStart(CStr(iKey(r, 1) & "")) = True
        End If
    Next r
    For r = 1 To UBound(key, 1)                                           ' the last backlog entry of each ticket without a start
        If Not IsError(key(r, 1)) And Not IsError(cat(r, 1)) Then
            k = CStr(key(r, 1) & "")
            If noStart.Exists(k) And CStr(cat(r, 1) & "") = "Backlog" And SerialOfCell(entered(r, 1), t) Then
                If Not last.Exists(k) Then
                    last(k) = t
                ElseIf t > last(k) Then
                    last(k) = t
                End If
            End If
        End If
    Next r
    For r = 1 To UBound(key, 1)
        If Not IsError(key(r, 1)) And Not IsError(isStart(r, 1)) Then
            k = CStr(key(r, 1) & "")
            If last.Exists(k) And Val(CStr(isStart(r, 1) & "")) = 1 And SerialOfCell(entered(r, 1), t) Then
                If Abs(t - last(k)) <= HALF_SECOND Then flagged(k) = True
            End If
        End If
    Next r
    If flagged.Count > 0 Then
        Cov "HLT65"
        SameSecondStarts = " | Started in the second of the backlog entry: " & flagged.Count & " tickets (no StartUtc, so left out of Flow Time and Flow Load)"
    End If
End Function

Private Function AsOfBeforeData(ByVal wb As Object) As String
    ' The status changes entered after AsOf: the open intervals and Flow Load stop at AsOf (flow/health.as_of_before_data). An AsOf that cannot be
    ' read is the business of the run, not of this field: nothing.
    Dim entered As Variant, asOf As Double, r As Long, n As Long, t As Double
    entered = ColumnValues(TableOf(wb, "StatusIntervals"), "EnteredUtc")
    If IsEmpty(entered) Or Not SheetExists(wb, "Config") Then Exit Function
    On Error Resume Next
    asOf = CDbl(AsOfFromConfig(wb.Worksheets(SheetName("Config"))))
    If Err.Number <> 0 Then Cov "HLT68": Err.Clear: Exit Function
    On Error GoTo 0
    For r = 1 To UBound(entered, 1)
        If SerialOfCell(entered(r, 1), t) Then
            If t > asOf + HALF_SECOND Then n = n + 1
        End If
    Next r
    If n > 0 Then
        Cov "HLT66"
        AsOfBeforeData = " | AsOf before the data: " & n & " status changes after AsOf (Flow Load stops at AsOf): set AsOf in " & SheetName("Config") & _
                         " to the time of the extraction"
    End If
End Function

Private Function VelocityCheckGap(ByVal wb As Object) As String
    ' The rows of the table VelocityCheck whose Gap is not 0 (an error counts): the Velocity does not add up from the status changes.
    Dim gap As Variant, r As Long, n As Long
    gap = ColumnValues(TableOf(wb, "VelocityCheck"), "Gap")
    If IsEmpty(gap) Then Exit Function
    For r = 1 To UBound(gap, 1)
        If IsError(gap(r, 1)) Then
            n = n + 1
        ElseIf Val(CStr(gap(r, 1) & "")) <> 0 Then
            n = n + 1
        End If
    Next r
    If n > 0 Then
        Cov "HLT67"
        VelocityCheckGap = " | VelocityCheck gap: " & n & " of " & UBound(gap, 1) & " rows (sheet " & SheetName("VelocityCheck") & ")"
    End If
End Function

Private Function ReconciliationFindings(ByVal wb As Object) As String
    ' The controls of the sheet Reconciliation to review: "Reconciliation: n of m controls to review", only when there is one. A row INFO (a point of
    ' attention, such as the tickets near a boundary) is neither OK nor to review: it is not counted.
    Dim res As Variant, r As Long, n As Long, m As Long
    res = ColumnValues(TableOf(wb, "Reconciliation"), "Result")
    If IsEmpty(res) Then Exit Function
    For r = 1 To UBound(res, 1)
        m = m + 1
        If IsError(res(r, 1)) Then
            n = n + 1
        ElseIf StrComp(CStr(res(r, 1) & ""), "REVIEW", vbBinaryCompare) = 0 Then
            n = n + 1
        End If
    Next r
    If n > 0 Then
        Cov "HLT62"
        ReconciliationFindings = " | Reconciliation: " & n & " of " & m & " controls to review (sheet " & SheetName("Reconciliation") & ")"
    End If
End Function

Private Function IsOtherTeam(ByVal team As Variant, ByVal r As Long) As Boolean
    ' True when row r of the column Team (a 2-D array, or Empty) is the team of the Program tickets that are not Features.
    If IsEmpty(team) Then Exit Function
    If IsError(team(r, 1)) Then Exit Function
    IsOtherTeam = (StrComp(TrimWhite(CStr(team(r, 1) & "")), OTHER_TEAM, vbTextCompare) = 0)
End Function

Private Function NotFeatures(ByVal wb As Object) As String
    ' The Program tickets that are not Features, and (when the ART has none) a FeatureTypes that matches nothing. Numbers and type names only.
    Dim team As Variant, typ As Variant, r As Long, nOther As Long, nArt As Long, seen As Object, names As String, n As Long, name As String, wanted As String
    team = ColumnValues(TableOf(wb, "ItemFlow"), "Team")
    If IsEmpty(team) Then Exit Function
    typ = ColumnValues(TableOf(wb, "ItemFlow"), "Type")
    Set seen = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(team, 1)
        If IsOtherTeam(team, r) Then
            nOther = nOther + 1
            If Not IsEmpty(typ) Then
                If Not IsError(typ(r, 1)) Then
                    name = TrimWhite(CStr(typ(r, 1) & ""))
                    If Len(name) > 0 And Not seen.Exists(name) Then
                        seen(name) = True
                        n = n + 1
                        If n <= MAX_NAMES Then names = names & IIf(Len(names) > 0, ", ", "") & name
                    End If
                End If
            End If
        ElseIf Not IsError(team(r, 1)) Then
            If StrComp(TrimWhite(CStr(team(r, 1) & "")), PROGRAM_TEAM, vbTextCompare) = 0 Then nArt = nArt + 1
        End If
    Next r
    If nOther = 0 Then Exit Function
    Cov "HLT59"
    NotFeatures = " | Not Features: " & nOther & " of " & (nOther + nArt) & " tickets of the Program (team " & OTHER_TEAM & " in " & SheetName("ItemFlow") & _
                  ", left out of the ART metrics)"
    If nArt = 0 Then
        Cov "HLT60"
        If SheetExists(wb, "Config") Then wanted = ConfigValue(wb.Worksheets(SheetName("Config")), "FeatureTypes")
        NotFeatures = NotFeatures & " | FeatureTypes """ & wanted & """ matches no ticket of the Program (types read: " & names & IIf(n > MAX_NAMES, ", ...", "") & _
                      "): type the name of the Feature type of Jira in " & SheetName("Config")
    End If
End Function

Private Function OutsideIterations(ByVal wb As Object) As String
    Dim done As Variant, starts As New Collection, ends As New Collection, row As Variant, s As Variant, e As Variant, team As Variant
    Dim r As Long, nDone As Long, nOut As Long
    If Not SheetExists(wb, "Iterations") Then Exit Function
    done = ColumnValues(TableOf(wb, "ItemFlow"), "DoneUtc")
    If IsEmpty(done) Then Exit Function
    team = ColumnValues(TableOf(wb, "ItemFlow"), "Team")                       ' the tickets "ART (other types)" feed no metric: not counted here
    For Each row In ReadTable(wb.Worksheets(SheetName("Iterations")))
        s = SerialOf(row("Start")): e = SerialOf(row("End"))
        If Not IsEmpty(s) And Not IsEmpty(e) Then starts.Add s: ends.Add e
    Next row
    For r = 1 To UBound(done, 1)
        If IsOtherTeam(team, r) Then
            Cov "HLT61"
        ElseIf IsDoneDate(done(r, 1)) Then
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
    For Each name In Array("StatusIntervals", "ItemFlow", "DailyLoad", "Metrics", "IterationFeatures", "PIMetrics", "PIStatusTime", "VelocityCheck", "FlowTimeCheck", "FlowEfficiencyCheck", "FlowLoadCheck", "PIIntervals", "DataChecks", "FillRates", "IssueLinks")
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

' ---------- The type of the columns written from Jira (python: flow/raw_types.py, the same rules and the same words) ----------

Public Function CellKind(ByVal v As Variant) As String
    ' empty | text | number | date | boolean | error | other: what a cell holds (raw_types.kind_of on the python side).
    If IsError(v) Then Cov "HLT08": CellKind = "error": Exit Function
    Select Case VarType(v)
        Case vbEmpty, vbNull
            Cov "HLT09": CellKind = "empty"
        Case vbString
            If Len(v) = 0 Then Cov "HLT10": CellKind = "empty" Else Cov "HLT11": CellKind = "text"
        Case vbBoolean
            Cov "HLT12": CellKind = "boolean"
        Case vbDate
            Cov "HLT13": CellKind = "date"
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            Cov "HLT14": CellKind = "number"
        Case Else
            Cov "HLT15": CellKind = "other"
    End Select
End Function

Private Function IsBlankCell(ByVal v As Variant) As Boolean
    ' A cell with nothing in it: Empty or "". An error is a value (it is the first thing to see).
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then IsBlankCell = True: Exit Function
    If VarType(v) = vbString Then IsBlankCell = (Len(v) = 0)
End Function

Private Function AllDigits(ByVal s As String) As Boolean
    AllDigits = Len(s) > 0
    If AllDigits Then AllDigits = Not (s Like "*[!0-9]*")
End Function

Public Function IsTicketKey(ByVal s As String) As Boolean
    ' ABC-12: a project key (a capital, then capitals, digits or _), a dash, digits.
    Dim p As Long, i As Long, c As Long
    p = InStrRev(s, "-")
    If p < 2 Then Cov "HLT16": Exit Function
    If Not AllDigits(Mid$(s, p + 1)) Then Cov "HLT17": Exit Function
    For i = 1 To p - 1
        c = AscW(Mid$(s, i, 1))
        If Not ((c >= 65 And c <= 90) Or (i > 1 And ((c >= 48 And c <= 57) Or c = 95))) Then Cov "HLT18": Exit Function
    Next i
    IsTicketKey = True
End Function

Private Function InClosedList(ByVal s As String, ByVal allowed As String) As Boolean
    Dim one As Variant
    For Each one In Split(allowed, ", ")
        If StrComp(s, CStr(one), vbBinaryCompare) = 0 Then InClosedList = True: Exit Function
    Next one
End Function

Private Function NoEmptyItem(ByVal s As String) As Boolean
    Dim one As Variant
    For Each one In Split(s, ",")
        If Len(TrimWhite(CStr(one))) = 0 Then Cov "HLT19": Exit Function
    Next one
    NoEmptyItem = True
End Function

Public Function FitsType(ByVal typ As String, ByVal v As Variant, ByVal allowed As String) As Boolean
    ' Is this (filled) value of the type? The same rules as raw_types.fits: a date is a real date cell from 2000 to 2099 (a text or a
    ' number, such as epoch milliseconds, is not), an id is digits kept as text, a key looks like ABC-12, a whole number has no decimals.
    ' (And does not short-circuit: a conversion of the value waits until its kind is known.)
    Dim kind As String, d As Double
    kind = CellKind(v)
    Select Case typ
        Case "datetime", "date"
            If kind <> "date" Then Cov "HLT56": Exit Function
            d = CDbl(v)
            If d < FIRST_SERIAL Or d >= LAST_SERIAL Then Cov "HLT57": Exit Function
            If typ = "date" And d - Int(d) > HALF_SECOND Then Cov "HLT58": Exit Function
            FitsType = True
        Case "id"
            Cov "HLT48"
            If kind = "text" Then FitsType = AllDigits(CStr(v))
        Case "key"
            Cov "HLT49"
            If kind = "text" Then FitsType = IsTicketKey(CStr(v))
        Case "integer"
            Cov "HLT50"
            If kind = "number" Then FitsType = (CDbl(v) = Int(CDbl(v)))
        Case "number"
            Cov "HLT51": FitsType = (kind = "number")
        Case "boolean"
            Cov "HLT52": FitsType = (kind = "boolean")
        Case "enum"
            Cov "HLT53"
            If kind = "text" Then FitsType = InClosedList(CStr(v), allowed)
        Case "list"
            Cov "HLT54"
            If kind = "text" Then FitsType = NoEmptyItem(CStr(v))
        Case Else
            Cov "HLT55": FitsType = (kind = "text")
    End Select
End Function

Private Function DigitsOf(ByVal v As Variant) As Long
    ' How many digits a whole number or a text of digits has: 13 = epoch milliseconds, 10 = epoch seconds; 0 for anything else.
    Dim kind As String
    kind = CellKind(v)
    If kind = "number" Then
        If CDbl(v) = Int(CDbl(v)) Then DigitsOf = Len(Format$(Abs(CDbl(v)), "0"))
    ElseIf kind = "text" Then
        If AllDigits(CStr(v)) Then DigitsOf = Len(CStr(v))
    End If
End Function

Private Function PlainNumber(ByVal d As Double) As String
    ' The text of a number as python writes it, in every locale: whole numbers in full, else 15 digits, a dot and a zero before it.
    Dim s As String
    If d = Int(d) Then PlainNumber = Format$(d, "0"): Exit Function
    s = Trim$(Str$(d))
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    PlainNumber = s
End Function

Private Function MaskedPart(ByVal text As String) As String
    ' ":9999AAA" - the shape of a text, never the text (DateShape masks digits and letters).
    MaskedPart = Mid$(DateShape(text), Len("String") + 1)
End Function

Public Function ShapeOfCell(ByVal v As Variant) As String
    ' "Double:9999999999999", "String:AAA-99", "Date:9999-99-99A99:99:99", "Boolean", "Error", "Empty": the type and the shape of a
    ' cell, none of its content (raw_types.shape_of).
    Dim d As Double
    Select Case CellKind(v)
        Case "empty": ShapeOfCell = "Empty"
        Case "other": ShapeOfCell = "Object"
        Case "boolean": ShapeOfCell = "Boolean"
        Case "error": ShapeOfCell = "Error"
        Case "date": ShapeOfCell = "Date" & MaskedPart(Format$(v, "yyyy-mm-dd") & "T" & Format$(v, "hh:nn:ss"))
        Case "number": ShapeOfCell = "Double" & MaskedPart(PlainNumber(CDbl(v)))
        Case Else: ShapeOfCell = "String" & MaskedPart(CStr(v))
    End Select
End Function

Public Function HintFor(ByVal typ As String, ByVal found As String, ByVal required As Boolean, ByVal allowed As String, ByVal digits As Long) As String
    ' What to look at, in one sentence, per (expected type, found kind): the same texts as raw_types.hint. No value in it.
    If found = "missing" Then
        Cov "HLT20": HintFor = "the column is not in the sheet: the sheet and RawTypesSpec differ (an old workbook?)": Exit Function
    End If
    If found = "empty" Then
        If typ = "datetime" Or typ = "date" Then
            Cov "HLT21"
            If required Then
                HintFor = "no cell holds a date: Jira's date was not read (see WARN J2D-DATA-003 and the INFO line 'changelog date key')"
            Else
                HintFor = "no ticket has this date: fine when the field is not used"
            End If
        Else
            Cov "HLT22"
            If required Then
                HintFor = "a required column is empty: the extraction wrote nothing there (check the field and its id in Config)"
            Else
                HintFor = "no ticket has a value: fine when the field is not configured or not used"
            End If
        End If
        Exit Function
    End If
    Select Case typ
        Case "datetime", "date"
            If digits >= 12 Then
                Cov "HLT23": HintFor = digits & " digits: epoch milliseconds, Jira sent a number and the date was not converted (modText.ParseJiraDate)"
            ElseIf digits >= 9 Then
                Cov "HLT24": HintFor = digits & " digits: epoch seconds, Jira sent a number and the date was not converted (modText.ParseJiraDate)"
            ElseIf found = "date" Then
                Cov "HLT25": HintFor = "a date outside 2000-2099, or with a time of day where Jira sends a day: a wrong unit or a typo in Jira"
            ElseIf found = "text" Then
                Cov "HLT26": HintFor = "a text that is not read as a date: Jira's format changed (see the shape; ParseJiraDate reads ISO and epoch)"
            Else
                Cov "HLT27": HintFor = "not a date: the cell holds another type than a date"
            End If
        Case "id", "key"
            If found = "number" Then
                Cov "HLT28": HintFor = "Excel turned the text into a number (leading apostrophe lost) or Jira sent a number: keep ids as text"
            Else
                Cov "HLT29": HintFor = "not a Jira id or key (ABC-12 or digits): the field is not the one expected"
            End If
        Case "integer", "number"
            If found = "text" Then
                Cov "HLT30": HintFor = "a text where a number is expected: Jira sent text, or the custom field is of another type (check its id in Config)"
            ElseIf found = "number" Then
                Cov "HLT31": HintFor = "a number with decimals where a whole number is expected"
            Else
                Cov "HLT32": HintFor = "not a number: the cell holds another type"
            End If
        Case "boolean"
            Cov "HLT33": HintFor = "not TRUE/FALSE: the cell holds another type"
        Case "enum"
            Cov "HLT34": HintFor = "not one of " & allowed & ": Jira added a value or the field is another one"
        Case "list"
            Cov "HLT35": HintFor = "an item of the list is empty: a separator without a value"
        Case Else
            Cov "HLT36": HintFor = "a text expected: Excel converted it (apostrophe lost) or Jira typed the value differently"
    End Select
End Function

Private Function KindName(ByVal slot As Long) As String
    KindName = Array("boolean", "date", "error", "number", "other", "text")(slot)
End Function

Private Function KindSlot(ByVal kind As String) As Long
    Select Case kind
        Case "boolean": KindSlot = 0
        Case "date": KindSlot = 1
        Case "error": KindSlot = 2
        Case "number": KindSlot = 3
        Case "other": KindSlot = 4
        Case Else: KindSlot = 5
    End Select
End Function

Private Function ProblemLines(ByVal rule As Object, ByVal found As String, ByVal cell As String, ByVal nRows As Long, ByVal filled As Long, _
                              ByVal kinds As String, ByVal shape As String, ByVal digits As Long, ByVal numFmt As String) As String
    ' The three lines of a problem, joined by vbLf (raw_types.TypeProblem.lines): the WARN line (code, column, types, source, first cell),
    ' the counts and the shape, the hint. Three because the log cuts a line at 300 characters. The cell format is the VBA side's addition.
    Dim head As String
    head = CStr(rule("Code")) & " | " & CStr(rule("Table")) & "." & CStr(rule("Column"))
    ProblemLines = head & " | expected " & CStr(rule("Type")) & " | found " & found & " | source " & CStr(rule("Source")) & _
        " | first filled " & IIf(Len(cell) > 0, cell, "-") & vbLf & _
        head & " | rows " & nRows & ", filled " & filled & ", empty " & (nRows - filled) & " | kinds in first " & TYPE_SAMPLE & ": " & _
        IIf(Len(kinds) > 0, kinds, "-") & " | value shape " & shape & IIf(Len(numFmt) > 0, " | cell format " & Left$(numFmt, 30), "") & vbLf & _
        head & " | hint: " & HintFor(CStr(rule("Type")), found, (CStr(rule("Required")) = "yes"), CStr(rule("Allowed")), digits)
End Function

Private Function TableShape(ByVal wb As Object, ByVal name As String, ByRef ws As Object, ByRef headerRow As Long, ByRef firstCol As Long, _
                            ByRef nCols As Long, ByRef nRows As Long) As Boolean
    ' Where the table `name` is and how big: a table of Excel (the flow tables) or, when the sheet has none, its plain range from A1 (the
    ' sheet Changelog is written as a range, not as a table). nRows is 0 when there is no row. True when there is at least one row.
    Dim lo As Object
    nRows = 0
    If Not SheetExists(wb, name) Then Cov "HLT45": Exit Function
    Set ws = wb.Worksheets(SheetName(name))
    Set lo = TableOf(wb, name)
    If Not lo Is Nothing Then
        headerRow = lo.HeaderRowRange.Row: firstCol = lo.Range.Column: nCols = lo.ListColumns.Count
        If Not lo.DataBodyRange Is Nothing Then nRows = lo.DataBodyRange.Rows.Count
    Else
        Cov "HLT46"
        headerRow = 1: firstCol = 1
        nCols = ws.Cells(1, ws.Columns.Count).End(-4159).Column       ' xlToLeft
        nRows = ws.Cells(ws.Rows.Count, 1).End(-4162).Row - 1         ' xlUp: the first column is the key, filled on every row
        If nRows < 0 Then nRows = 0
    End If
    If nRows = 1 Then
        If IsBlankCell(ws.Cells(headerRow + 1, firstCol).Value) Then Cov "HLT47": nRows = 0     ' a table keeps one blank row when it has none
    End If
    TableShape = (nRows > 0)
End Function

Private Function ColumnIndex(ByVal ws As Object, ByVal headerRow As Long, ByVal firstCol As Long, ByVal nCols As Long, ByVal header As String) As Long
    ' The column (in the sheet) whose header is `header`, 0 when there is none.
    Dim c As Long
    For c = firstCol To firstCol + nCols - 1
        If CStr(ws.Cells(headerRow, c).Value) = header Then ColumnIndex = c: Exit Function
    Next c
End Function

Private Function BlockValues(ByVal ws As Object, ByVal row1 As Long, ByVal row2 As Long, ByVal col As Long) As Variant
    ' The cells of a column from row1 to row2 as a 2-D array (rows, 1): a one-cell Value is not an array, so it is wrapped.
    Dim one(1 To 1, 1 To 1) As Variant
    If row1 = row2 Then
        one(1, 1) = ws.Cells(row1, col).Value
        BlockValues = one
    Else
        BlockValues = ws.Range(ws.Cells(row1, col), ws.Cells(row2, col)).Value
    End If
End Function

Private Function ColumnProblem(ByVal ws As Object, ByVal headerRow As Long, ByVal firstCol As Long, ByVal nCols As Long, ByVal nRows As Long, _
                               ByVal rule As Object, ByRef filled As Long, ByRef found As String) As String
    ' The three log lines (ProblemLines) of the column when its first filled cell is not of its type (or it is empty and required, or it is
    ' missing); "" when all is well. `filled` is the number of cells with something in them (CountA: one call, whatever the rows), `found`
    ' the kind of the first one. Only the first TYPE_SAMPLE filled cells are read, by blocks, so that 1 000 000 rows cost almost nothing.
    Dim idx As Long, vals As Variant, r As Long, r0 As Long, r1 As Long, got As Long, firstRow As Long, first As Variant, typ As String
    Dim allowed As String, counts(0 To 5) As Long, kinds As String, i As Long, cell As Object
    typ = CStr(rule("Type")): allowed = CStr(rule("Allowed"))
    filled = 0: found = ""
    idx = ColumnIndex(ws, headerRow, firstCol, nCols, CStr(rule("Column")))
    If idx = 0 Then
        Cov "HLT37": found = "missing"
        ColumnProblem = ProblemLines(rule, found, "", nRows, 0, "", "Empty", 0, "")
        Exit Function
    End If
    filled = CLng(Application.WorksheetFunction.CountA(ws.Range(ws.Cells(headerRow + 1, idx), ws.Cells(headerRow + nRows, idx))))
    r0 = 1
    Do While filled > 0 And r0 <= nRows And got < TYPE_SAMPLE
        r1 = r0 + BLOCK_ROWS - 1
        If r1 > nRows Then r1 = nRows
        vals = BlockValues(ws, headerRow + r0, headerRow + r1, idx)
        For r = 1 To r1 - r0 + 1
            If Not IsBlankCell(vals(r, 1)) Then
                got = got + 1
                If got = 1 Then firstRow = r0 + r - 1: first = vals(r, 1)
                i = KindSlot(CellKind(vals(r, 1))): counts(i) = counts(i) + 1
                If got = TYPE_SAMPLE Then Exit For
            End If
        Next r
        r0 = r1 + 1
    Loop
    If got = 0 Then
        filled = 0: found = "empty"
        If CStr(rule("Required")) = "yes" Then
            Cov "HLT38": ColumnProblem = ProblemLines(rule, found, "", nRows, 0, "", "Empty", 0, "")
        End If
        Exit Function
    End If
    found = CellKind(first)
    If FitsType(typ, first, allowed) Then Cov "HLT39": Exit Function
    For i = 0 To 5
        If counts(i) > 0 Then kinds = kinds & IIf(Len(kinds) > 0, ", ", "") & KindName(i) & " x " & counts(i)
    Next i
    Set cell = ws.Cells(headerRow + firstRow, idx)
    ColumnProblem = ProblemLines(rule, found, cell.Address(False, False), nRows, filled, kinds, ShapeOfCell(first), DigitsOf(first), CStr(cell.NumberFormat))
End Function

Private Sub WriteFills(ByVal pending As Collection, ByVal fillsLine As String, ByVal fillsCount As Long)
    ' The fill lines of the table just read, after its problems (the order of raw_types.log_lines).
    If fillsCount > 0 Then pending.Add fillsLine
    Do While pending.Count > 0
        LogLine "INFO", CStr(pending(1))
        pending.Remove 1
    Loop
End Sub

Public Function TypeProblems(ByVal wb As Object) As String
    ' The first filled cell of each column written from Jira must be of the type its label promises (the sheet RawTypesSpec, built from
    ' raw_types.py): a datetime column holds a date, an id column digits as text... Never stops the run (a WARN line per problem, with its
    ' support code J2D-DATA-004..013 and what helps to debug; per table the INFO line of the filled cells of each column; one summary).
    ' Returns the field " | Column types: n column(s), first Changelog.CreatedUtc expected datetime found empty; ..." or "".
    Dim rule As Variant, table As String, ws As Object, headerRow As Long, firstCol As Long, nCols As Long, tableRows As Long
    Dim filled As Long, found As String, line As String, parts As Variant, begun As Boolean, item As String
    Dim tables As Long, columns As Long, checked As Long, problems As Long, names As String
    Dim fillsLine As String, fillsCount As Long, pending As New Collection
    On Error GoTo Failed
    If Not SheetExists(wb, "RawTypesSpec") Then Cov "HLT07": Exit Function
    LogBegin "column types": begun = True
    For Each rule In ReadTable(wb.Worksheets(SheetName("RawTypesSpec")))
        If CStr(rule("Table")) <> table Then
            WriteFills pending, fillsLine, fillsCount
            table = CStr(rule("Table")): fillsCount = 0
            If TableShape(wb, table, ws, headerRow, firstCol, nCols, tableRows) Then tables = tables + 1
            fillsLine = "types " & table & " | rows " & tableRows & " | filled: "
        End If
        If tableRows = 0 Then
            Cov "HLT40"
        Else
            columns = columns + 1
            line = ColumnProblem(ws, headerRow, firstCol, nCols, tableRows, rule, filled, found)
            If found <> "missing" Then
                item = CStr(rule("Column")) & " " & filled
                If fillsCount > 0 And Len(fillsLine) + 2 + Len(item) > FILL_LINE_MAX Then
                    Cov "HLT44": pending.Add fillsLine
                    fillsLine = "types " & table & " | filled (continued): ": fillsCount = 0
                End If
                fillsLine = fillsLine & IIf(fillsCount > 0, ", ", "") & item: fillsCount = fillsCount + 1
            End If
            If filled > 0 Then checked = checked + 1
            If Len(line) > 0 Then
                Cov "HLT41"
                problems = problems + 1
                parts = Split(line, vbLf)
                LogLine "WARN", parts(0)
                LogLine "INFO", parts(1)
                LogLine "INFO", parts(2)
                If problems <= MAX_NAMES Then
                    names = names & IIf(Len(names) > 0, "; ", "") & table & "." & CStr(rule("Column")) & " expected " & CStr(rule("Type")) & " found " & found
                End If
            End If
        End If
    Next rule
    WriteFills pending, fillsLine, fillsCount
    LogLine "INFO", "column types | tables " & tables & " | columns " & columns & " | checked " & checked & " | problems " & problems
    LogEnd "column types"
    If problems > 0 Then
        Cov "HLT42"
        TypeProblems = " | Column types: " & problems & " column(s), first " & names & IIf(problems > MAX_NAMES, "; ...", "")
    End If
    Exit Function
Failed:
    Cov "HLT43"
    LogLine "WARN", "column types: the check failed | error " & Err.Number
    If begun Then LogEnd "column types"
    TypeProblems = ""
End Function
