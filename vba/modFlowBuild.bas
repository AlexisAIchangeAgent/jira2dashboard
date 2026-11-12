Attribute VB_Name = "modFlowBuild"
Option Explicit
' Builds the flow sheets (ItemFlow, DailyLoad, Metrics, Controls) and the dashboard (ChartData, the charts) after the
' extraction. Nothing is computed here: Excel formulas do it. The formulas and the description of the charts are NOT in
' this code: they are rows of the sheets FlowFormulas (Table | Column | Formula) and ChartSpec, written by the Python
' generator from the tested formula text. An empty Formula marks a column of values (raw data or keys written here).
' The sizes follow the data: Issues, StatusIntervals, ItemFlow = one row per extracted row, DailyLoad = one row per day
' and team, Metrics = one row per iteration and team. Every call rebuilds everything, so no row of an older run stays.

Private Const XL_SRC_RANGE As Long = 1
Private Const XL_YES As Long = 1
Private Const XL_CALC_MANUAL As Long = -4135
Private Const XL_CALC_AUTOMATIC As Long = -4105
Private Const CHARTS_SHEET As String = "Charts"      ' the sheet of the charts, made at build (flow/dashboard_spec.CHARTS_SHEET)
Private Const BAND_MARKER As Long = -4115             ' xlMarkerStyleDash, a long bar (flow/dashboard_spec.BAND_MARKER)

Public Function SheetExists(ByVal wb As Object, ByVal name As String) As Boolean
    Cov "FLB01"
    Dim ws As Object
    For Each ws In wb.Worksheets
        If ws.Name = name Then SheetExists = True: Exit Function
    Next ws
End Function

Private Function AddSheet(ByVal wb As Object, ByVal name As String) As Object
    Cov "FLB02"
    Set AddSheet = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
    AddSheet.Name = name
End Function

Private Function AsText(ByVal s As String) As String
    ' A leading apostrophe keeps Excel from reading "I1" or "2026-01" as something else than text.
    AsText = "'" & s
End Function

Private Sub UnlistAll(ByVal ws As Object)
    Do While ws.ListObjects.Count > 0
        ws.ListObjects(1).Unlist
    Loop
End Sub

Public Sub ClearFlowSheets(ByVal wb As Object)
    ' Back to "nothing extracted": derived sheets and charts deleted, raw sheets emptied (the Dashboard stays).
    Cov "FLB03"
    Dim nm As Variant
    DeleteDerived wb
    For Each nm In Array("Issues", "StatusIntervals", "IssueLinks")
        If SheetExists(wb, CStr(nm)) Then
            UnlistAll wb.Worksheets(CStr(nm))
            wb.Worksheets(CStr(nm)).Cells.Clear
        End If
    Next nm
End Sub

Private Sub DeleteDerived(ByVal wb As Object)
    Cov "FLB04"
    Dim nm As Variant, i As Long, alerts As Boolean
    If SheetExists(wb, CHARTS_SHEET) Then                 ' the charts and the notes under them; the sheet and its selector stay
        For i = wb.Worksheets(CHARTS_SHEET).ChartObjects.Count To 1 Step -1
            wb.Worksheets(CHARTS_SHEET).ChartObjects(i).Delete
        Next i
        wb.Worksheets(CHARTS_SHEET).Range("A4:R400").ClearContents
    End If
    alerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error GoTo Restore
    For Each nm In Array("ChartData", "ItemFlow", "DailyLoad", "Metrics", "PIMetrics", "DataChecks", "FillRates", "Controls")
        If SheetExists(wb, CStr(nm)) Then wb.Worksheets(CStr(nm)).Delete
    Next nm
Restore:
    Application.DisplayAlerts = alerts
    If Err.Number <> 0 Then Err.Raise Err.Number, "DeleteDerived", Err.Description
End Sub

Private Function ReadFormulas(ByVal wb As Object) As Object
    ' table name -> Collection of Array(column, formula), in the order of the sheet FlowFormulas.
    Cov "FLB05"
    Dim out As Object, row As Variant, name As String
    Set out = CreateObject("Scripting.Dictionary")
    For Each row In ReadTable(wb.Worksheets("FlowFormulas"))
        name = CStr(row("Table"))
        If Not out.Exists(name) Then Set out(name) = New Collection
        out(name).Add Array(CStr(row("Column")), CStr(row("Formula") & ""))
    Next row
    Set ReadFormulas = out
End Function

Private Function SettingOf(ByVal spec As Object, ByVal table As String, ByVal column As String) As String
    Cov "FLB06"
    Dim pair As Variant
    For Each pair In spec(table)
        If pair(0) = column Then SettingOf = pair(1): Exit Function
    Next pair
End Function

Private Function TeamNames(ByVal wsProjects As Object, ByVal programTeam As String) As Collection
    ' The Program reports under its own team name (ART), then the teams of the Team projects, each once.
    Cov "FLB07"
    Dim out As New Collection, seen As Object, row As Variant, team As String
    Set seen = CreateObject("Scripting.Dictionary")
    out.Add programTeam
    seen(programTeam) = True
    For Each row In ReadTable(wsProjects)
        team = TrimWhite(CStr(row("Team") & ""))
        If StrComp(TrimWhite(CStr(row("Role") & "")), "Team", vbTextCompare) = 0 And Len(team) > 0 Then
            If Not seen.Exists(team) Then Cov "FLB08": seen(team) = True: out.Add team
        End If
    Next row
    Set TeamNames = out
End Function

Private Function IterationRows(ByVal wsIterations As Object) As Collection
    ' Iteration-label | Start | End -> Collection of Array(name, start serial, end serial)
    Cov "FLB09"
    Dim out As New Collection, row As Variant, s As Variant, e As Variant
    For Each row In ReadTable(wsIterations)
        s = SerialOf(row("Start")): e = SerialOf(row("End"))
        If IsEmpty(s) Or IsEmpty(e) Then
            Cov "FLB28"
            Err.Raise vbObjectError + 701, "BuildFlowSheets", "Sheet Iterations: Start and End of iteration " & row("Iteration-label") & " must be dates."
        End If
        If e <= s Then
            Cov "FLB27"
            Err.Raise vbObjectError + 704, "BuildFlowSheets", "Sheet Iterations: iteration " & row("Iteration-label") & " ends before it starts."
        End If
        out.Add Array(CStr(row("Iteration-label")), s, e)
    Next row
    Set IterationRows = out
End Function

Private Function SerialOf(ByVal v As Variant) As Variant
    ' A date serial from a date cell, a number, or ISO text (2026-06-01). Empty if it is none of them.
    Cov "FLB29"
    Dim m As Variant
    If VarType(v) = vbDate Then SerialOf = CDbl(v): Exit Function
    If VarType(v) = vbString Then
        m = ParseMoment(CStr(v))
        If Not IsEmpty(m) Then SerialOf = CDbl(m)
        Exit Function
    End If
    If IsNumeric(v) And Not IsEmpty(v) Then SerialOf = CDbl(v)
End Function

Private Function CreateTable(ByVal ws As Object, ByVal name As String, ByVal cols As Collection, ByVal dataRows As Long) As Object
    ' Headers are written, then the table is made over header + data rows (at least one data row).
    Cov "FLB10"
    Dim i As Long, last As Long, headers() As Variant
    ReDim headers(1 To 1, 1 To cols.Count)
    For i = 1 To cols.Count
        headers(1, i) = cols(i)(0)
    Next i
    ws.Range(ws.Cells(1, 1), ws.Cells(1, cols.Count)).Value = headers          ' one write, not one call per cell
    last = dataRows
    If last < 1 Then last = 1
    Set CreateTable = ws.ListObjects.Add(XL_SRC_RANGE, ws.Range(ws.Cells(1, 1), ws.Cells(last + 1, cols.Count)), , XL_YES)
    CreateTable.Name = name
End Function

Private Sub FillFormulas(ByVal lo As Object, ByVal cols As Collection)
    ' One formula per column of the table: Excel turns it into the formula of every row (calculated column).
    Cov "FLB11"
    Dim i As Long, num As Long, desc As String
    For i = 1 To cols.Count
        If Len(cols(i)(1)) > 0 Then
            On Error Resume Next
            lo.ListColumns(CStr(cols(i)(0))).DataBodyRange.Formula2 = cols(i)(1)
            num = Err.Number: desc = Err.Description: Err.Clear
            On Error GoTo 0
            If num <> 0 Then
                Err.Raise vbObjectError + 707, "BuildFlowSheets", "Excel refused the formula of " & lo.Name & "." & cols(i)(0) & ": " & desc
            End If
        End If
    Next i
End Sub

Private Function UsedRows(ByVal ws As Object) As Long
    UsedRows = ws.Cells(ws.Rows.Count, 1).End(-4162).Row - 1            ' -4162 = xlUp
End Function

Private Sub PrepareRaw(ByVal ws As Object, ByVal cols As Collection)
    ' The extraction wrote plain values in the first columns; drop an older table and its helper columns.
    Cov "FLB12"
    Dim rawColumns As Long, i As Long
    rawColumns = RawCount(cols)
    UnlistAll ws
    ws.Range(ws.Columns(rawColumns + 1), ws.Columns(rawColumns + 40)).Clear
    If Len(CStr(ws.Cells(1, 1).Value)) = 0 Then
        Err.Raise vbObjectError + 702, "BuildFlowSheets", "Sheet " & ws.Name & " is empty: extract the data from Jira first."
    End If
    For i = 1 To rawColumns                                  ' the extraction and FlowFormulas must agree on the columns
        If CStr(ws.Cells(1, i).Value) <> cols(i)(0) Then
            Err.Raise vbObjectError + 706, "BuildFlowSheets", "Sheet " & ws.Name & ": column " & i & " is '" & ws.Cells(1, i).Value & _
                      "' but FlowFormulas expects '" & cols(i)(0) & "'. Rebuild the workbook with scripts\run.ps1 build."
        End If
    Next i
End Sub

Private Function RawCount(ByVal cols As Collection) As Long
    Cov "FLB13"
    Dim i As Long
    For i = 1 To cols.Count
        If Len(cols(i)(1)) = 0 Then RawCount = RawCount + 1
    Next i
End Function

Private Sub RequireDynamicArrays()
    ' The formulas of the plan use LET, XLOOKUP and FILTER (Excel 365 or 2021): say so before anything is read or cleared.
    Cov "FLB35"
    If IsError(Application.Evaluate("LET(x,1,x)")) Then
        Err.Raise vbObjectError + 708, "BuildFlowSheets", "This workbook needs Excel 365 or Excel 2021 (dynamic array formulas: LET, XLOOKUP, FILTER)."
    End If
End Sub

Private Sub LoadPlan(ByVal wb As Object, ByRef spec As Object, ByRef teams As Collection, ByRef iterations As Collection)
    ' Everything the build needs from the sheets, checked: the errors come before anything is changed.
    Dim nm As Variant
    RequireDynamicArrays
    For Each nm In Array("FlowFormulas", "ChartSpec", "ChecksSpec", "Config", "Iterations", "PIs", "Projects", "Workflows", _
                         "Transitions", "Dashboard", CHARTS_SHEET, "Issues", "StatusIntervals", "IssueLinks")
        If Not SheetExists(wb, CStr(nm)) Then
            Err.Raise vbObjectError + 705, "BuildFlowSheets", "Sheet " & nm & " is missing: rebuild the workbook with scripts\run.ps1 build."
        End If
    Next nm
    Set spec = ReadFormulas(wb)
    Set teams = TeamNames(wb.Worksheets("Projects"), SettingOf(spec, "Dashboard", "ProgramTeam"))
    Set iterations = IterationRows(wb.Worksheets("Iterations"))
    If iterations.Count = 0 Then
        Cov "FLB15"
        Err.Raise vbObjectError + 703, "BuildFlowSheets", "Sheet Iterations has no iteration: fill Iteration-label, Start and End first."
    End If
End Sub

Public Sub ValidateBuildInputs(ByVal wb As Object)
    ' Called before the extraction: a bad Iterations sheet or an old Excel must not be found after ten minutes of reading Jira,
    ' with the earlier results already cleared.
    Dim spec As Object, teams As Collection, iterations As Collection
    LoadPlan wb, spec, teams, iterations
End Sub

Public Sub BuildFlowSheetsIn(ByVal wb As Object)
    Cov "FLB14"
    Dim spec As Object, teams As Collection, iterations As Collection, calcMode As Long, screen As Boolean
    Dim num As Long, desc As String
    LoadPlan wb, spec, teams, iterations
    calcMode = Application.Calculation
    screen = Application.ScreenUpdating
    On Error GoTo Failed                                              ' from here Failed gives both back
    Application.ScreenUpdating = False
    Application.Calculation = XL_CALC_MANUAL
    BuildAll wb, spec, teams, iterations
    LogBegin "flow sheets: calculation"
    Application.Calculation = calcMode                   ' back to automatic: Excel calculates everything now
    If calcMode <> XL_CALC_AUTOMATIC Then Cov "FLB36": Application.CalculateFull      ' Excel left in manual mode by the user
    Application.ScreenUpdating = screen
    LogEnd "flow sheets: calculation"
    On Error Resume Next
    wb.Worksheets("Dashboard").Activate                  ' the sheets added by the build must not leave another one in front
    Exit Sub
Failed:
    num = Err.Number: desc = Err.Description
    Application.Calculation = calcMode
    Application.ScreenUpdating = screen
    Err.Raise num, "BuildFlowSheets", desc
End Sub

Private Sub BuildAll(ByVal wb As Object, ByVal spec As Object, ByVal teams As Collection, ByVal iterations As Collection)
    Cov "FLB16"
    Dim wsI As Object, wsS As Object, wsF As Object, wsD As Object, wsM As Object, wsC As Object, wsL As Object, wsR As Object
    Dim loI As Object, loS As Object, loF As Object, loD As Object, loM As Object, loC As Object, loL As Object, loR As Object
    Dim wsP As Object, loP As Object, pis As Collection
    LogBegin "flow sheets: prepare"
    DeleteDerived wb
    Set wsI = wb.Worksheets("Issues"): Set wsS = wb.Worksheets("StatusIntervals")
    PrepareRaw wsI, spec("Issues")
    PrepareRaw wsS, spec("StatusIntervals")
    Set wsL = wb.Worksheets("IssueLinks")
    PrepareRaw wsL, spec("IssueLinks")
    Dim issues As Long, checks As Long, fills As Long
    issues = UsedRows(wsI)
    BuildControls wb, spec
    LogEnd "flow sheets: prepare"
    LogBegin "flow sheets: keys"
    Set wsF = AddSheet(wb, "ItemFlow"): Set wsD = AddSheet(wb, "DailyLoad"): Set wsM = AddSheet(wb, "Metrics")
    Set wsC = AddSheet(wb, "DataChecks"): Set wsR = AddSheet(wb, "FillRates"): Set wsP = AddSheet(wb, "PIMetrics")
    Set pis = PiNames(wb, SettingOf(spec, "Dashboard", "NoPI"))
    WriteDailyKeys wsD, iterations, teams
    WriteMetricKeys wsM, iterations, teams
    WritePiKeys wsP, pis
    checks = WriteCheckKeys(wsC, wb.Worksheets("ChecksSpec"))
    fills = WriteFillKeys(wsR, spec("FillFields"), wb, teams)
    LogEnd "flow sheets: keys"
    ' every table exists before any formula is written: the formulas refer to each other's columns
    LogBegin "flow sheets: tables"
    Set loI = CreateTable(wsI, "Issues", spec("Issues"), issues)
    Set loS = CreateTable(wsS, "StatusIntervals", spec("StatusIntervals"), UsedRows(wsS))
    Set loF = CreateTable(wsF, "ItemFlow", spec("ItemFlow"), issues)
    Set loD = CreateTable(wsD, "DailyLoad", spec("DailyLoad"), DaysOf(iterations) * teams.Count)
    Set loM = CreateTable(wsM, "Metrics", spec("Metrics"), iterations.Count * teams.Count)
    Set loC = CreateTable(wsC, "DataChecks", spec("DataChecks"), checks)
    Set loL = CreateTable(wsL, "IssueLinks", spec("IssueLinks"), UsedRows(wsL))
    Set loR = CreateTable(wsR, "FillRates", spec("FillRates"), fills)
    Set loP = CreateTable(wsP, "PIMetrics", spec("PIMetrics"), pis.Count)
    LogEnd "flow sheets: tables"
    LogBegin "flow sheets: formulas StatusIntervals"
    FillFormulas loS, spec("StatusIntervals")
    LogEnd "flow sheets: formulas StatusIntervals"
    LogBegin "flow sheets: formulas ItemFlow"
    FillFormulas loF, spec("ItemFlow")
    LogEnd "flow sheets: formulas ItemFlow"
    LogBegin "flow sheets: formulas IssueLinks"
    FillFormulas loL, spec("IssueLinks")
    LogEnd "flow sheets: formulas IssueLinks"
    LogBegin "flow sheets: formulas FillRates"
    FillFormulas loR, spec("FillRates")
    LogEnd "flow sheets: formulas FillRates"
    LogBegin "flow sheets: formulas DataChecks"
    FillFormulas loC, spec("DataChecks")
    LogEnd "flow sheets: formulas DataChecks"
    LogBegin "flow sheets: formulas DailyLoad"
    FillFormulas loD, spec("DailyLoad")
    LogEnd "flow sheets: formulas DailyLoad"
    LogBegin "flow sheets: formulas Metrics"
    FillFormulas loM, spec("Metrics")
    LogEnd "flow sheets: formulas Metrics"
    LogBegin "flow sheets: formulas PIMetrics"
    FillFormulas loP, spec("PIMetrics")
    LogEnd "flow sheets: formulas PIMetrics"
    wsD.Columns(1).NumberFormat = "yyyy-mm-dd"
    wsM.Range(wsM.Columns(2), wsM.Columns(3)).NumberFormat = "yyyy-mm-dd"
    wsP.Range(wsP.Columns(2), wsP.Columns(3)).NumberFormat = "yyyy-mm-dd"
    wsP.Columns(6).NumberFormat = "0%"
    LogBegin "flow sheets: charts"
    BuildCharts wb, spec, teams, iterations, pis
    LogEnd "flow sheets: charts"
End Sub

Private Sub BuildControls(ByVal wb As Object, ByVal spec As Object)
    ' The as-of date and the Flow Load mode, read from Config: cells with names that the formulas use.
    Cov "FLB17"
    Dim ws As Object, i As Long
    Set ws = AddSheet(wb, "Controls")
    For i = 1 To spec("Controls").Count
        ws.Cells(i, 1).Value = spec("Controls")(i)(0)
        ws.Cells(i, 2).Formula2 = spec("Controls")(i)(1)
        wb.Names.Add spec("Controls")(i)(0), "=Controls!$B$" & i
    Next i
    ws.Cells(1, 2).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Columns("A:B").AutoFit
End Sub

Private Function DaysOf(ByVal iterations As Collection) As Long
    ' Whole days from the first iteration start to the last iteration end (end excluded).
    Cov "FLB18"
    Dim p As Variant, first As Double, last As Double
    first = Int(iterations(1)(1)): last = Int(iterations(1)(2))
    For Each p In iterations
        If Int(p(1)) < first Then first = Int(p(1))
        If Int(p(2)) > last Then last = Int(p(2))
    Next p
    DaysOf = CLng(last - first)
End Function

Private Sub WriteDailyKeys(ByVal ws As Object, ByVal iterations As Collection, ByVal teams As Collection)
    ' Date | Team, one row per day and team. Values, not formulas: a formula typed in a table column fills the whole column.
    Cov "FLB19"
    Dim p As Variant, first As Double, n As Long, d As Long, t As Long, r As Long, data() As Variant
    first = Int(iterations(1)(1))
    For Each p In iterations
        If Int(p(1)) < first Then first = Int(p(1))
    Next p
    n = DaysOf(iterations) * teams.Count
    If n = 0 Then Exit Sub
    ReDim data(1 To n, 1 To 2)
    For d = 0 To DaysOf(iterations) - 1
        For t = 1 To teams.Count
            r = r + 1
            data(r, 1) = first + d
            data(r, 2) = AsText(CStr(teams(t)))
        Next t
    Next d
    ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, 2)).Value = data
End Sub

Private Sub WritePiKeys(ByVal ws As Object, ByVal pis As Collection)
    ' PI, one row per PI of PiNames (the sheet PIs, then the PIs of Features it lacks, or the one row noPi): the dates and the
    ' figures are the formulas of the table PIMetrics (flow/formulas.PIMETRICS_COLUMNS).
    Cov "FLB40"
    Dim r As Long, data() As Variant
    ReDim data(1 To pis.Count, 1 To 1)
    For r = 1 To pis.Count
        data(r, 1) = AsText(CStr(pis(r)))
    Next r
    ws.Range(ws.Cells(2, 1), ws.Cells(pis.Count + 1, 1)).Value = data
End Sub

Private Sub WriteMetricKeys(ByVal ws As Object, ByVal iterations As Collection, ByVal teams As Collection)
    ' Iteration-label | IterationStart | IterationEnd | Team, one row per iteration and team.
    Cov "FLB20"
    Dim p As Variant, t As Long, r As Long, data() As Variant
    ReDim data(1 To iterations.Count * teams.Count, 1 To 4)
    For Each p In iterations
        For t = 1 To teams.Count
            r = r + 1
            data(r, 1) = AsText(CStr(p(0))): data(r, 2) = p(1): data(r, 3) = p(2): data(r, 4) = AsText(CStr(teams(t)))
        Next t
    Next p
    ws.Range(ws.Cells(2, 1), ws.Cells(r + 1, 4)).Value = data
End Sub

Private Function WriteCheckKeys(ByVal ws As Object, ByVal wsSpec As Object) As Long
    ' Id | Title | Severity, one row per data check of the sheet ChecksSpec. Returns the number of checks.
    Cov "FLB30"
    Dim rows As Collection, row As Variant, r As Long, data() As Variant, names As Variant, c As Long
    names = Array("Id", "Title", "Severity", "Need1", "Need2", "Need3", "Source")
    Set rows = ReadTable(wsSpec)
    WriteCheckKeys = rows.Count
    If rows.Count = 0 Then Exit Function
    ReDim data(1 To rows.Count, 1 To 7)
    For Each row In rows
        r = r + 1
        For c = 1 To 7
            data(r, c) = AsText(CStr(row(names(c - 1)) & ""))
        Next c
    Next row
    ws.Range(ws.Cells(2, 1), ws.Cells(rows.Count + 1, 7)).Value = data
End Function

Private Function IssueTypes(ByVal wsIssues As Object) As Collection
    ' The distinct values of the column Type of the sheet Issues, in order of appearance.
    Cov "FLB32"
    Dim out As New Collection, seen As Object, last As Long, data As Variant, r As Long
    Set seen = CreateObject("Scripting.Dictionary")
    Set IssueTypes = out
    last = UsedRows(wsIssues)
    If last < 1 Then Exit Function
    data = IssueColumn(wsIssues, "Type", last)
    For r = 1 To last
        AddType out, seen, CStr(data(r, 1) & "")
    Next r
End Function

Private Function IssueColumn(ByVal wsIssues As Object, ByVal header As String, ByVal last As Long) As Variant
    ' Rows 2 to last + 1 of the column of the sheet Issues whose header (row 1) is header, always as an array
    ' (1 To last, 1 To 1): a one-cell range does not come back as an array.
    Dim col As Long, lastCol As Long, one(1 To 1, 1 To 1) As Variant
    lastCol = wsIssues.Cells(1, wsIssues.Columns.Count).End(-4159).Column      ' -4159 = xlToLeft
    For col = 1 To lastCol
        If CStr(wsIssues.Cells(1, col).Value) = header Then Exit For
    Next col
    If col > lastCol Then Err.Raise vbObjectError + 709, "IssueColumn", "The sheet Issues has no column " & header & ": run the extraction again."
    If last = 1 Then
        one(1, 1) = wsIssues.Cells(2, col).Value
        IssueColumn = one
    Else
        IssueColumn = wsIssues.Range(wsIssues.Cells(2, col), wsIssues.Cells(last + 1, col)).Value
    End If
End Function

Private Sub AddType(ByVal out As Collection, ByVal seen As Object, ByVal v As String)
    If Len(v) = 0 Then Exit Sub
    If Not seen.Exists(v) Then seen(v) = True: out.Add v
End Sub

Private Function WriteFillKeys(ByVal ws As Object, ByVal fields As Collection, ByVal wb As Object, _
                               ByVal teams As Collection) As Long
    ' Field | Team | Type, one row per field x team x issue type. A field that needs a Jira field id is left out while that
    ' id is empty in Config (a system field always counts). Returns the number of rows.
    Cov "FLB31"
    Dim types As Collection, names As New Collection, pair As Variant, t As Long, k As Long, f As Long, r As Long, data() As Variant
    Set types = IssueTypes(wb.Worksheets("Issues"))
    For Each pair In fields
        If Len(CStr(pair(1))) = 0 Then
            names.Add CStr(pair(0))
        ElseIf Len(ConfigValue(wb.Worksheets("Config"), CStr(pair(1)))) > 0 Then
            names.Add CStr(pair(0))
        End If
    Next pair
    WriteFillKeys = names.Count * teams.Count * types.Count
    If WriteFillKeys = 0 Then Exit Function
    ReDim data(1 To WriteFillKeys, 1 To 3)
    For f = 1 To names.Count
        For t = 1 To teams.Count
            For k = 1 To types.Count
                r = r + 1
                data(r, 1) = AsText(names(f)): data(r, 2) = AsText(CStr(teams(t))): data(r, 3) = AsText(types(k))
            Next k
        Next t
    Next f
    ws.Range(ws.Cells(2, 1), ws.Cells(WriteFillKeys + 1, 3)).Value = data
End Function

' --- the dashboard: ChartData (one block of formulas per chart) and the charts -----------------------------------

Private Sub BuildCharts(ByVal wb As Object, ByVal spec As Object, ByVal teams As Collection, ByVal iterations As Collection, _
                        ByVal pis As Collection)
    Cov "FLB21"
    Dim data As Object, dash As Object, row As Variant, index As Long, stride As Long
    Set dash = wb.Worksheets(CHARTS_SHEET)               ' a sheet of their own: on the Dashboard they hid its cells
    Set data = AddSheet(wb, "ChartData")
    stride = IIf(pis.Count > iterations.Count, pis.Count, iterations.Count)     ' a block: one row per iteration or PI
    WriteSelectorList wb, dash, data, spec, teams
    For Each row In ReadTable(wb.Worksheets("ChartSpec"))
        BuildChart dash, data, row, 1 + index * (stride + 4), teams, iterations, pis
        index = index + 1
    Next row
End Sub

Private Function PiNames(ByVal wb As Object, ByVal noPi As String) As Collection
    ' The PIs of the sheet PIs (first column, from row 2), in its order, then those that Features carry but the sheet lacks
    ' (FeaturePis); noPi alone when there is none, so that the chart still has a row (a gap). They are the keys of the table
    ' PIMetrics and the axis of the Feature Predictability chart. The dates are not read here: the table looks them up itself,
    ' and a PI it cannot read is not computable, never an error of the build. Reference: flow/metrics.feature_predictability.
    Dim out As New Collection, seen As Object, ws As Object, last As Long, r As Long, v As String
    Cov "FLB37"
    Set seen = CreateObject("Scripting.Dictionary")
    Set ws = wb.Worksheets("PIs")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row     ' -4162 = xlUp
    For r = 2 To last
        v = Trim$(CStr(ws.Cells(r, 1).Value & ""))
        If Len(v) > 0 Then
            out.Add v
            seen(UCase$(v)) = True                       ' MATCH and COUNTIFS of the table ignore the case
        End If
    Next r
    FeaturePis out, seen, wb
    If out.Count = 0 Then Cov "FLB38": out.Add noPi
    Set PiNames = out
End Function

Private Sub FeaturePis(ByVal out As Collection, ByVal seen As Object, ByVal wb As Object)
    ' Adds to out the PIs of the Features (type in FeatureTypes, case-sensitive like IsFeature) of the sheet Issues that the
    ' sheet PIs lacks (AL, 2026-11-12), by year then number: without dates they are not computable, but the RTE sees them.
    ' Only a valid PI (PI#XX-YYYY) counts: a malformed one is for DC17 and DC29.
    Dim wsI As Object, last As Long, types As Variant, pis As Variant, features As Object, extra As New Collection
    Dim r As Long, j As Long, v As String, t As Variant
    Set wsI = wb.Worksheets("Issues")
    last = UsedRows(wsI)
    If last < 1 Then Exit Sub
    Set features = CreateObject("Scripting.Dictionary")              ' binary compare: case-sensitive, as FIND in IsFeature
    For Each t In Split(ConfigValue(wb.Worksheets("Config"), "FeatureTypes"), ",")
        If Len(TrimWhite(CStr(t))) > 0 Then features(TrimWhite(CStr(t))) = True
    Next t
    types = IssueColumn(wsI, "Type", last)
    pis = IssueColumn(wsI, "PI", last)
    For r = 1 To last
        v = CStr(pis(r, 1) & "")
        If IsPiName(v) Then
            If features.Exists(CStr(types(r, 1) & "")) And Not seen.Exists(UCase$(v)) Then
                Cov "FLB41"
                seen(UCase$(v)) = True
                For j = 1 To extra.Count
                    If PiOrder(v) < PiOrder(CStr(extra(j))) Then Exit For
                Next j
                If j > extra.Count Then
                    extra.Add v
                Else
                    Cov "FLB42"                                      ' before a later PI already found
                    extra.Add v, , j
                End If
            End If
        End If
    Next r
    For j = 1 To extra.Count
        out.Add extra(j)
    Next j
End Sub

Private Function IsPiName(ByVal v As String) As Boolean
    ' PI#XX-YYYY, XX and YYYY digits, PI# in capitals: the format of DC17 and DC29 (flow/checks.valid_pi).
    Dim k As Variant
    If Len(v) <> 10 Or Left$(v, 3) <> "PI#" Or Mid$(v, 6, 1) <> "-" Then Exit Function
    For Each k In Array(4, 5, 7, 8, 9, 10)
        If InStr(1, "0123456789", Mid$(v, k, 1), vbBinaryCompare) = 0 Then Exit Function
    Next k
    IsPiName = True
End Function

Private Function PiOrder(ByVal v As String) As String
    PiOrder = Right$(v, 4) & Mid$(v, 4, 2)                             ' PI#XX-YYYY: by year, then number
End Function

Private Sub WriteSelectorList(ByVal wb As Object, ByVal dash As Object, ByVal data As Object, ByVal spec As Object, _
                              ByVal teams As Collection)
    ' The team of the Distribution chart is picked from a list: All teams, then the teams.
    Cov "FLB22"
    Dim i As Long, cell As Object, allTeams As String, found As Boolean, listColumn As Long
    allTeams = SettingOf(spec, "Dashboard", "AllTeams")
    Set cell = dash.Range(SettingOf(spec, "Dashboard", "SelectorCell"))
    listColumn = teams.Count + 4                       ' right of the widest chart block (one column per team, or 5 work types)
    If listColumn < 10 Then listColumn = 10
    data.Cells(1, listColumn).Value = "Teams for the selector"
    data.Cells(2, listColumn).Value = AsText(allTeams)
    found = (CStr(cell.Value) = allTeams)
    For i = 1 To teams.Count
        data.Cells(2 + i, listColumn).Value = AsText(CStr(teams(i)))
        If CStr(cell.Value) = CStr(teams(i)) Then found = True
    Next i
    cell.Validation.Delete
    cell.Validation.Add 3, 1, 1, "=ChartData!" & data.Range(data.Cells(2, listColumn), data.Cells(teams.Count + 2, listColumn)).Address   ' list, stop on error, between
    If Not found Then Cov "FLB23": cell.Value = allTeams
End Sub

Private Sub BuildChart(ByVal dash As Object, ByVal data As Object, ByVal row As Variant, ByVal top As Long, _
                       ByVal teams As Collection, ByVal iterations As Collection, ByVal pis As Collection)
    Cov "FLB24"
    Dim names() As String, count As Long, i As Long, j As Long, first As Long, last As Long, point As String
    Dim area As Object, shape As Object, ch As Object, s As Object, rng As Object
    Dim rowsCount As Long, firstDay As Double, p As Variant, band() As String, isDay As Boolean, axis() As Variant
    Dim isPi As Boolean
    If Len(CStr(row("SeriesNames"))) > 0 Then
        names = Split(CStr(row("SeriesNames")), ",")
    Else
        ReDim names(0 To teams.Count - 1)
        For i = 1 To teams.Count: names(i - 1) = teams(i): Next i
    End If
    count = UBound(names) + 1
    isDay = (CStr(row("Axis")) = "day")                      ' one row per day (daily curve), per PI or per iteration
    isPi = (CStr(row("Axis")) = "pi")
    rowsCount = iterations.Count
    If isDay Then rowsCount = DaysOf(iterations)
    If isPi Then rowsCount = pis.Count
    first = top + 2: last = first + rowsCount - 1
    ReDim axis(1 To rowsCount, 1 To 1)
    data.Cells(top, 1).Value = CStr(row("Title"))
    data.Cells(top, 1).Font.Bold = True
    If isDay Then
        Cov "FLB33"
        data.Cells(top + 1, 1).Value = "Day"
        firstDay = Int(iterations(1)(1))
        For Each p In iterations
            If Int(p(1)) < firstDay Then firstDay = Int(p(1))
        Next p
        For i = 1 To rowsCount
            axis(i, 1) = firstDay + i - 1
        Next i
        data.Range(data.Cells(first, 1), data.Cells(last, 1)).Value = axis
        data.Range(data.Cells(first, 1), data.Cells(last, 1)).NumberFormat = "yyyy-mm-dd"
    ElseIf isPi Then
        Cov "FLB39"
        data.Cells(top + 1, 1).Value = "PI"
        For i = 1 To rowsCount
            axis(i, 1) = AsText(CStr(pis(i)))
        Next i
        data.Range(data.Cells(first, 1), data.Cells(last, 1)).Value = axis
    Else
        data.Cells(top + 1, 1).Value = "Iteration-label"
        For i = 1 To rowsCount
            axis(i, 1) = AsText(CStr(iterations(i)(0)))
        Next i
        data.Range(data.Cells(first, 1), data.Cells(last, 1)).Value = axis
    End If
    For j = 1 To count
        data.Cells(top + 1, 1 + j).Value = AsText(names(j - 1))
    Next j
    point = Replace(CStr(row("PointFormula")), "{iteration}", "$A" & first)
    If InStr(point, "{series}") > 0 Then
        Cov "FLB25"                                          ' one formula per series: the name is part of the text
        For j = 1 To count
            data.Range(data.Cells(first, 1 + j), data.Cells(last, 1 + j)).Formula2 = Replace(point, "{series}", names(j - 1))
        Next j
    Else
        data.Range(data.Cells(first, 2), data.Cells(last, 1 + count)).Formula2 = Replace(point, "{name}", "B$" & (top + 1))
    End If
    Set area = dash.Range(dash.Cells(CLng(row("AnchorRow")), CLng(row("AnchorColumn"))), _
                          dash.Cells(CLng(row("AnchorRow")) + CLng(row("Rows")) - 1, CLng(row("AnchorColumn")) + CLng(row("Columns")) - 1))
    Set shape = dash.Shapes.AddChart2(-1, CLng(row("ChartType")), area.Left, area.Top, area.Width, area.Height, True)
    shape.Name = CStr(row("Name"))
    Set ch = shape.Chart
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop
    For j = 1 To count
        Set s = ch.SeriesCollection.NewSeries
        s.Name = "=ChartData!" & data.Cells(top + 1, 1 + j).Address
        s.Values = data.Range(data.Cells(first, 1 + j), data.Cells(last, 1 + j))
        s.XValues = data.Range(data.Cells(first, 1), data.Cells(last, 1))
    Next j
    If Len(CStr(row("Band"))) > 0 Then
        Cov "FLB34"                                          ' a target band: two constant series, dashed, after the data
        band = Split(CStr(row("Band")), ";")
        For j = 0 To 1
            data.Cells(top + 1, 2 + count + j).Value = IIf(j = 0, "Target band low", "Target band high")
            data.Range(data.Cells(first, 2 + count + j), data.Cells(last, 2 + count + j)).Value = Val(band(j))
            Set s = ch.SeriesCollection.NewSeries
            s.Name = "=ChartData!" & data.Cells(top + 1, 2 + count + j).Address
            s.Values = data.Range(data.Cells(first, 2 + count + j), data.Cells(last, 2 + count + j))
            s.XValues = data.Range(data.Cells(first, 1), data.Cells(last, 1))
            s.ChartType = 65                                 ' xlLineMarkers: a line needs two points, one PI shows only its markers
            s.Format.Line.DashStyle = 4                      ' msoLineDash
            s.Format.Line.ForeColor.RGB = RGB(128, 128, 128)
            s.MarkerStyle = BAND_MARKER                      ' after the line format, which would recolour the markers
            s.MarkerSize = 12
            s.MarkerForegroundColor = RGB(128, 128, 128)
            s.MarkerBackgroundColor = RGB(128, 128, 128)
        Next j
    End If
    ch.HasTitle = True
    ch.ChartTitle.Text = CStr(row("Title"))
    ch.HasLegend = True
    ch.Legend.Position = -4107                               ' xlLegendPositionBottom
    ch.DisplayBlanksAs = 1                                   ' xlNotPlotted: a figure that is not computable is a gap, never 0
    If Len(CStr(row("AxisFormat"))) > 0 Then Cov "FLB26": ch.Axes(2).TickLabels.NumberFormat = CStr(row("AxisFormat"))
    Set rng = data.Range(data.Cells(first, 2), data.Cells(last, 1 + count))
    dash.Cells(CLng(row("AnchorRow")) + CLng(row("Rows")), CLng(row("AnchorColumn"))).Formula2 = _
        Replace(CStr(row("NoteFormula")), "{range}", "ChartData!" & rng.Address)
End Sub
