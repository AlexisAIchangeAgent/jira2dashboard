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
Private Const ITERATION_CHARTS_SHEET As String = "IterationCharts"   ' the charts of the Features per iteration (dashboard_spec.ITERATION_CHARTS_SHEET)
Private Const STATUSES_TOKEN As String = "{statuses}"   ' SeriesNames of the chart Cycle time by status: its series are StatusNames (dashboard_spec.STATUSES_TOKEN)
Private Const NO_DEPTH As Long = 1000000                  ' flow/pi_status_time.NO_DEPTH: a status that no transition reaches comes last
Private Const NO_STATUS As String = "(no status)"        ' the one status of the table PIStatusTime when the Feature workflow has no Active or Wait status (dashboard_spec.NO_STATUS)
Private Const BAND_MARKER As Long = -4115            ' xlMarkerStyleDash, a long bar (flow/dashboard_spec.BAND_MARKER)

Public Function SheetExists(ByVal wb As Object, ByVal name As String) As Boolean
    Cov "FLB01"
    Dim ws As Object
    For Each ws In wb.Worksheets
        If ws.Name = SheetName(name) Then SheetExists = True: Exit Function
    Next ws
End Function

Private Function AddSheet(ByVal wb As Object, ByVal name As String) As Object
    Cov "FLB02"
    Set AddSheet = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
    AddSheet.Name = SheetName(name)
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
    For Each nm In Array("Issues", "StatusIntervals", "IssueLinks", "Changelog", "PIIntervals")
        If SheetExists(wb, CStr(nm)) Then
            UnlistAll wb.Worksheets(SheetName(CStr(nm)))
            wb.Worksheets(SheetName(CStr(nm))).Cells.Clear
        End If
    Next nm
End Sub

Private Sub DeleteDerived(ByVal wb As Object)
    ' Every result of a build: the charts, then the tables (the charts read the tables).
    Cov "FLB04"
    DeleteCharts wb
    DeleteTables wb
End Sub

Private Sub DeleteCharts(ByVal wb As Object)
    ' The charts, the notes under them and the sheet ChartData; the chart sheets and their selector stay.
    Dim nm As Variant, i As Long
    For Each nm In Array(CHARTS_SHEET, ITERATION_CHARTS_SHEET)
        If SheetExists(wb, CStr(nm)) Then
            For i = wb.Worksheets(SheetName(CStr(nm))).ChartObjects.Count To 1 Step -1
                wb.Worksheets(SheetName(CStr(nm))).ChartObjects(i).Delete
            Next i
            wb.Worksheets(SheetName(CStr(nm))).Range("A4:R400").ClearContents
        End If
    Next nm
    DeleteSheets wb, Array("ChartData")
End Sub

Private Sub DeleteTables(ByVal wb As Object)
    ' The calculated sheets (tables and formulas); the raw sheets stay.
    DeleteSheets wb, Array("ItemFlow", "DailyLoad", "Metrics", "IterationFeatures", "PIMetrics", "PIStatusTime", "VelocityCheck", "FlowTimeCheck", "FlowEfficiencyCheck", "FlowLoadCheck", "FlowDistCheck", "MetricItems", "Reconciliation", _
                           "DataChecks", "FillRates", "Controls")
End Sub

Private Sub DeleteSheets(ByVal wb As Object, ByVal names As Variant)
    Dim nm As Variant, alerts As Boolean
    alerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error GoTo Restore
    For Each nm In names
        If SheetExists(wb, CStr(nm)) Then wb.Worksheets(SheetName(CStr(nm))).Delete
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
    For Each row In ReadTable(wb.Worksheets(SheetName("FlowFormulas")))
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

Public Function TeamNames(ByVal wsProjects As Object, ByVal programTeam As String) As Collection
    ' The Program reports under its own team name (ART), then the teams of the Team projects, each once. A Team project whose cell Team
    ' is empty is its own team (the project key).
    Cov "FLB07"
    Dim out As New Collection, seen As Object, row As Variant, team As String
    Set seen = CreateObject("Scripting.Dictionary")
    out.Add programTeam
    seen(programTeam) = True
    For Each row In ReadTable(wsProjects)
        team = TrimWhite(CStr(row("Team") & ""))
        If Len(team) = 0 Then team = TrimWhite(CStr(row("Project") & ""))
        If StrComp(TrimWhite(CStr(row("Role") & "")), "Team", vbTextCompare) = 0 And Len(team) > 0 Then
            If Not seen.Exists(team) Then Cov "FLB08": seen(team) = True: out.Add team
        End If
    Next row
    Set TeamNames = out
End Function

Public Function StatusNames(ByVal wb As Object) As Collection
    ' The stack of the chart Cycle time by status and the status keys of the table PIStatusTime: the statuses that the Features (the ART: type in FeatureTypes,
    ' project of Role Program) visited and whose category is Active or Wait, in the order of the normal flow: by depth in the transitions of the Jira workflow of
    ' the Feature types (JiraIssueTypes, JiraWorkflows, JiraTransitions; a status without depth comes last), then by the earliest Seq at which a Feature visited it,
    ' then by name (binary). One placeholder when there is none, so that the chart and the table still have a row (a gap). Reference: flow/pi_status_time.feature_statuses.
    ' It reads the raw sheets Issues and StatusIntervals (the formulas of the build are not calculated yet).
    Cov "FLB50"
    Dim out As New Collection, depth As Object, first As Object, program As Object, features As Object, wsI As Object, wsS As Object, lastI As Long, lastS As Long
    Dim keys As Variant, kinds As Variant, projects As Variant, sKey As Variant, sStatus As Variant, sCat As Variant, sSeq As Variant, isFeature As Object
    Dim names() As String, n As Long, i As Long, j As Long, r As Long, t As Variant, row As Variant, status As String, cat As String, seq As Long, tmp As String
    Set first = CreateObject("Scripting.Dictionary")
    Set isFeature = CreateObject("Scripting.Dictionary")
    Set program = CreateObject("Scripting.Dictionary")
    Set features = CreateObject("Scripting.Dictionary")                      ' binary compare: case-sensitive, as FIND in IsFeature
    For Each row In ReadTable(wb.Worksheets(SheetName("Projects")))
        If StrComp(TrimWhite(CStr(row("Role") & "")), "Program", vbTextCompare) = 0 Then program(TrimWhite(CStr(row("Project")))) = True
    Next row
    For Each t In Split(ConfigValue(wb.Worksheets(SheetName("Config")), "FeatureTypes"), ",")
        If Len(TrimWhite(CStr(t))) > 0 Then features(TrimWhite(CStr(t))) = True
    Next t
    Set wsI = wb.Worksheets(SheetName("Issues")): Set wsS = wb.Worksheets(SheetName("StatusIntervals"))
    lastI = UsedRows(wsI): lastS = UsedRows(wsS)
    If lastI >= 1 And lastS >= 1 Then
        keys = IssueColumn(wsI, "Key", lastI): kinds = IssueColumn(wsI, "Type", lastI): projects = IssueColumn(wsI, "Project", lastI)
        For r = 1 To lastI
            If program.Exists(CStr(projects(r, 1) & "")) And features.Exists(CStr(kinds(r, 1) & "")) Then isFeature(CStr(keys(r, 1))) = True
        Next r
        sKey = IssueColumn(wsS, "Key", lastS): sStatus = IssueColumn(wsS, "Status", lastS): sCat = IssueColumn(wsS, "Category", lastS): sSeq = IssueColumn(wsS, "Seq", lastS)
        For r = 1 To lastS
            cat = CStr(sCat(r, 1) & "")
            If (cat = "Active" Or cat = "Wait") And isFeature.Exists(CStr(sKey(r, 1))) Then
                status = CStr(sStatus(r, 1) & ""): seq = CLng(Val(CStr(sSeq(r, 1))))
                If Not first.Exists(status) Then
                    Cov "FLB58": first(status) = seq
                ElseIf seq < first(status) Then
                    Cov "FLB61": first(status) = seq                          ' the earliest position at which a Feature visited it
                End If
            End If
        Next r
    End If
    Set depth = FeatureDepths(wb, program, features)
    n = first.Count
    If n > 0 Then
        ReDim names(1 To n)
        For Each t In first.Keys
            i = i + 1: names(i) = CStr(t)
        Next t
        For i = 2 To n                                                       ' insertion sort on (depth, earliest Seq, name): the order of Python's sorted
            tmp = names(i): j = i - 1
            Do While j >= 1
                If StatusBefore(names(j), tmp, depth, first) Then Exit Do
                names(j + 1) = names(j): j = j - 1
            Loop
            names(j + 1) = tmp
        Next i
        For i = 1 To n
            out.Add names(i)
        Next i
    End If
    If out.Count = 0 Then Cov "FLB51": out.Add NO_STATUS
    Set StatusNames = out
End Function

Private Function StatusBefore(ByVal a As String, ByVal b As String, ByVal depth As Object, ByVal first As Object) As Boolean
    ' True when status a comes before or with b: depth first (none = last), then the earliest Seq, then the name in binary order.
    Dim da As Long, db As Long
    da = NO_DEPTH: db = NO_DEPTH
    If depth.Exists(a) Then da = depth(a)
    If depth.Exists(b) Then db = depth(b)
    If da <> db Then StatusBefore = (da < db): Exit Function
    If first(a) <> first(b) Then StatusBefore = (first(a) < first(b)): Exit Function
    StatusBefore = (StrComp(a, b, vbBinaryCompare) <= 0)
End Function

Private Function FeatureDepths(ByVal wb As Object, ByVal program As Object, ByVal features As Object) As Object
    ' status -> its smallest depth (moves from the creation status) in the Jira workflows that the Feature types of the Program projects follow
    ' (flow/jira_workflows.feature_depths). Empty when the sheets of the Jira configuration are not there (Read the Jira configuration was not run).
    Dim out As Object, names As Object, row As Variant, wf As Variant, initial As String, moves As Collection, d As Object, s As Variant
    Dim flowRows As Collection, moveRows As Collection
    Set out = CreateObject("Scripting.Dictionary")
    Set FeatureDepths = out
    If Not (SheetExists(wb, "JiraIssueTypes") And SheetExists(wb, "JiraWorkflows") And SheetExists(wb, "JiraTransitions")) Then Exit Function
    Set names = CreateObject("Scripting.Dictionary")
    For Each row In ReadTable(wb.Worksheets(SheetName("JiraIssueTypes")))
        If program.Exists(CStr(row("Project") & "")) And features.Exists(CStr(row("IssueType") & "")) And Len(CStr(row("Workflow") & "")) > 0 Then
            Cov "FLB59": names(CStr(row("Workflow"))) = True
        End If
    Next row
    Set flowRows = ReadTable(wb.Worksheets(SheetName("JiraWorkflows")))
    Set moveRows = ReadTable(wb.Worksheets(SheetName("JiraTransitions")))
    For Each wf In names.Keys
        initial = "": Set moves = New Collection
        For Each row In flowRows
            If CStr(row("Workflow") & "") = CStr(wf) And Val(CStr(row("Initial") & "")) = 1 Then initial = CStr(row("Status") & "")
        Next row
        For Each row In moveRows
            If CStr(row("Workflow") & "") = CStr(wf) Then moves.Add Array(CStr(row("From") & ""), CStr(row("To") & ""))
        Next row
        Set d = FlowDepths(initial, moves)
        For Each s In d.Keys
            If Not out.Exists(CStr(s)) Then
                Cov "FLB60": out(CStr(s)) = d(s)
            ElseIf d(s) < out(CStr(s)) Then
                Cov "FLB62": out(CStr(s)) = d(s)                              ' the smallest depth of the workflows of the Feature types
            End If
        Next s
    Next wf
End Function

Public Function SelectorColumn(ByVal teamCount As Long, ByVal featureTeamCount As Long, ByVal statusCount As Long) As Long
    ' The column of ChartData that holds the list of teams of the selector: right of the widest chart block (the axis, then one column
    ' per team, per team of the Features, per work type (5) or per status).
    Dim wide As Long
    wide = teamCount
    If featureTeamCount > wide Then Cov "FLB55": wide = featureTeamCount
    If statusCount > wide Then Cov "FLB53": wide = statusCount
    SelectorColumn = wide + 4
    If SelectorColumn < 10 Then SelectorColumn = 10
End Function

Private Sub WritePiStatusKeys(ByVal ws As Object, ByVal pis As Collection, ByVal statuses As Collection)
    ' PI | Status, one row per PI of PiNames and status of StatusNames; the dates and the figures are the formulas of the table
    ' PIStatusTime (flow/formulas.PISTATUSTIME_COLUMNS).
    Cov "FLB54"
    Dim p As Long, s As Long, r As Long, data() As Variant
    ReDim data(1 To pis.Count * statuses.Count, 1 To 2)
    For p = 1 To pis.Count
        For s = 1 To statuses.Count
            r = r + 1
            data(r, 1) = AsText(CStr(pis(p))): data(r, 2) = AsText(CStr(statuses(s)))
        Next s
    Next p
    ws.Range(ws.Cells(2, 1), ws.Cells(r + 1, 2)).Value = data
End Sub

Private Function IterationRows(ByVal wsIterations As Object) As Collection
    ' Iteration-label | Start | End -> Collection of Array(label, start serial, end serial). The columns are found by their header (RequireColumns
    ' checked the three) and a row is read when one of its three cells is filled, as the Python path reads it (params_book.sheet_from_grid): a row
    ' with dates and no label no longer disappears without a word (ReadTable skips a row whose first cell is empty), and the notes of the column
    ' Example are not read. The label is the key of every table by iteration (Metrics, the check tables, IterationFeatures, the axis of the
    ' charts): trimmed as the Python path trims it, never empty, never twice, case ignored as SUMIFS, XLOOKUP and COUNTIFS compare (AL, 2026-11-26).
    Cov "FLB09"
    Dim out As New Collection, cols As Object, col As Variant, last As Long, r As Long, s As Variant, e As Variant
    Dim lv As Variant, sv As Variant, ev As Variant, label As String, seen As Object
    Set IterationRows = out
    Set seen = CreateObject("Scripting.Dictionary")
    Set cols = HeaderColumns(wsIterations)
    For Each col In Array(cols("Iteration-label"), cols("Start"), cols("End"))
        r = wsIterations.Cells(wsIterations.Rows.Count, CLng(col)).End(-4162).Row     ' -4162 = xlUp
        If r > last Then last = r
    Next col
    For r = 2 To last
        lv = wsIterations.Cells(r, cols("Iteration-label")).Value
        sv = wsIterations.Cells(r, cols("Start")).Value: ev = wsIterations.Cells(r, cols("End")).Value
        If Not (IsBlankCell(lv) And IsBlankCell(sv) And IsBlankCell(ev)) Then
            If IsError(lv) Then label = "" Else label = TrimWhite(CStr(lv & ""))
            If Len(label) = 0 Then
                Cov "FLB72"
                Err.Raise vbObjectError + 714, "BuildFlowSheets", "Sheet Iterations: row " & r & " has no Iteration-label: give each iteration its own label."
            End If
            s = SerialOf(sv): e = SerialOf(ev)
            If Not IsEmpty(s) Then If s < 36526 Or s >= 73051 Then s = Empty   ' 2000 to 2099: 1026 for 2026, or 20261005 typed as a number, is a typo
            If Not IsEmpty(e) Then If e < 36526 Or e >= 73051 Then e = Empty
            If IsEmpty(s) Or IsEmpty(e) Then
                Cov "FLB28"
                Err.Raise vbObjectError + 701, "BuildFlowSheets", "Sheet Iterations: Start and End of iteration " & label & " must be dates from 2000 to 2099."
            End If
            If s <> Int(s) Or e <> Int(e) Then                          ' a time of day: the daily load counts whole days (the Python path refuses it too)
                Cov "FLB63"
                Err.Raise vbObjectError + 701, "BuildFlowSheets", "Sheet Iterations: Start and End of iteration " & label & _
                          " are whole days (no time of day): the daily load counts days."
            End If
            If e <= s Then
                Cov "FLB27"
                Err.Raise vbObjectError + 704, "BuildFlowSheets", "Sheet Iterations: iteration " & label & " ends before it starts."
            End If
            If seen.Exists(UCase$(label)) Then
                Cov "FLB73"
                Err.Raise vbObjectError + 714, "BuildFlowSheets", "Sheet Iterations: iteration " & label & " is listed twice (case ignored, as Excel " & _
                          "compares): give each iteration its own label."
            End If
            seen(UCase$(label)) = True
            out.Add Array(label, s, e)
        End If
    Next r
End Function

Private Function IsBlankCell(ByVal v As Variant) As Boolean
    ' Empty, or blanks only; an error value is not blank.
    If Not IsError(v) Then IsBlankCell = (Len(TrimWhite(CStr(v & ""))) = 0)
End Function

Private Function HeaderColumns(ByVal ws As Object) As Object
    ' The headers of row 1, trimmed and case kept as the Python path reads them (params_book.sheet_from_grid) -> their column; the first one wins.
    Dim c As Long, v As Variant, heading As String, out As Object
    Set out = CreateObject("Scripting.Dictionary")
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(-4159).Column        ' -4159 = xlToLeft
        v = ws.Cells(1, c).Value
        If Not IsError(v) Then
            heading = TrimWhite(CStr(v & ""))
            If Len(heading) > 0 And Not out.Exists(heading) Then out(heading) = c
        End If
    Next c
    Set HeaderColumns = out
End Function

Public Function SerialOf(ByVal v As Variant) As Variant
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

Private Sub BuildMetricItems(ByVal wb As Object, ByVal spec As Object)
    ' The sheet MetricItems: the lists of the tickets behind the figures (flow/trace.metric_items), four blocks side by side with a blank column
    ' between. Row 1 the title of a block, row 2 the names of its columns, row 3 one spilled formula per column: Excel fills the rows below by itself.
    ' A column with no formula is filled by the spill of the first one (the 4th list: formulas.METRICITEMS_AT_END).
    Cov "FLB56"
    Dim ws As Object, blocks As Variant, b As Long, cols As Collection, titles As Collection, i As Long, c As Long, num As Long, desc As String, nm As String
    Set ws = wb.Worksheets(SheetName("MetricItems"))
    Set titles = spec("MetricItemsTitle")
    blocks = Array("MetricItemsDone", "MetricItemsLeftOut", "MetricItemsCutoff", "MetricItemsAtEnd")
    c = 1
    For b = 0 To UBound(blocks)
        Set cols = spec(CStr(blocks(b)))
        ws.Cells(1, c).Value = titles(b + 1)(0)
        ws.Cells(1, c).Font.Bold = True
        For i = 1 To cols.Count
            nm = CStr(cols(i)(0))
            ws.Cells(2, c + i - 1).Value = nm
            ws.Cells(2, c + i - 1).Font.Bold = True
            num = 0
            If Len(CStr(cols(i)(1))) > 0 Then
                On Error Resume Next
                ws.Cells(3, c + i - 1).Formula2 = cols(i)(1)
                num = Err.Number: desc = Err.Description: Err.Clear
                On Error GoTo 0
            End If
            If num <> 0 Then Err.Raise vbObjectError + 707, "BuildFlowSheets", "Excel refused the formula of MetricItems." & blocks(b) & "." & nm & ": " & desc
            If nm = "Start" Or nm = "End" Or nm = "DoneUtc" Then ws.Columns(c + i - 1).NumberFormat = "yyyy-mm-dd hh:mm"
        Next i
        c = c + cols.Count + 1
    Next b
    ws.Range(ws.Cells(2, 1), ws.Cells(60, c)).Columns.AutoFit                  ' the headers and the first rows: not the titles of row 1, not every spilled row
End Sub

Private Sub WriteOneFormula(ByVal cell As Object, ByVal formula As String, ByVal what As String)
    ' One formula of a cell, with the error of the build when Excel refuses it.
    Dim num As Long, desc As String
    On Error Resume Next
    cell.Formula2 = formula
    num = Err.Number: desc = Err.Description: Err.Clear
    On Error GoTo 0
    If num <> 0 Then Err.Raise vbObjectError + 707, "BuildFlowSheets", "Excel refused the formula of " & what & ": " & desc
End Sub

Private Sub BuildReconciliation(ByVal wb As Object, ByVal spec As Object)
    ' The sheet Reconciliation (flow/reconciliation.py): the table of controls A1:F13 (Expected and Found are a formula of their own in each row),
    ' the census of the types (spilled formulas from H1) and, below the table, the parameters in force that the figures read.
    Cov "FLB57"
    Dim ws As Object, lo As Object, keys As Collection, expected As Collection, found As Collection, cols As Collection, census As Collection
    Dim params As Collection, unknown As Collection, ended As Collection, i As Long, top As Long
    Set ws = wb.Worksheets(SheetName("Reconciliation"))
    Set keys = spec("ReconciliationControl"): Set expected = spec("ReconciliationExpected"): Set found = spec("ReconciliationFound")
    Set cols = spec("Reconciliation"): Set census = spec("Census"): Set params = spec("ParametersInForce"): Set unknown = spec("Unclassified"): Set ended = spec("DoneStatuses")
    ' the cells are written BEFORE the table exists: a formula typed in one cell of a table column fills the whole column
    For i = 1 To keys.Count
        ws.Cells(i + 1, 1).Value = keys(i)(0)
        ws.Cells(i + 1, 2).Value = keys(i)(1)
        WriteOneFormula ws.Cells(i + 1, 3), CStr(expected(i)(1)), "Reconciliation.Expected of " & keys(i)(0)
        WriteOneFormula ws.Cells(i + 1, 4), CStr(found(i)(1)), "Reconciliation.Found of " & keys(i)(0)
    Next i
    Set lo = CreateTable(ws, "Reconciliation", cols, keys.Count)
    FillFormulas lo, cols                                                       ' Difference and Result: the same formula in every row
    For i = 1 To census.Count
        ws.Cells(1, 7 + i).Value = census(i)(0)
        ws.Cells(1, 7 + i).Font.Bold = True
        WriteOneFormula ws.Cells(2, 7 + i), CStr(census(i)(1)), "Census." & census(i)(0)
    Next i
    For i = 1 To unknown.Count                                                  ' the statuses that no rule classifies, from column O (15)
        ws.Cells(1, 14 + i).Value = unknown(i)(0)
        ws.Cells(1, 14 + i).Font.Bold = True
        WriteOneFormula ws.Cells(2, 14 + i), CStr(unknown(i)(1)), "Unclassified." & unknown(i)(0)
    Next i
    For i = 1 To ended.Count                                                    ' the statuses that end the flow and the tickets in them, from column S (19)
        ws.Cells(1, 18 + i).Value = ended(i)(0)
        ws.Cells(1, 18 + i).Font.Bold = True
        WriteOneFormula ws.Cells(2, 18 + i), CStr(ended(i)(1)), "DoneStatuses." & ended(i)(0)
    Next i
    top = keys.Count + 4
    ws.Cells(top - 1, 1).Value = "Parameters in force (what the figures read: compare them with Jira)"
    ws.Cells(top - 1, 1).Font.Bold = True
    For i = 1 To params.Count
        ws.Cells(top + i - 1, 1).Value = params(i)(0)
        WriteOneFormula ws.Cells(top + i - 1, 2), CStr(params(i)(1)), "ParametersInForce." & params(i)(0)
        If InStr(1, CStr(params(i)(0)), "iteration", vbTextCompare) > 0 Or InStr(1, CStr(params(i)(0)), "AsOf", vbTextCompare) > 0 Then
            ws.Cells(top + i - 1, 2).NumberFormat = "yyyy-mm-dd hh:mm"           ' the labels of the dates (formulas.PARAMETERS_IN_FORCE, pinned by test_blueprint)
        End If
    Next i
    ws.Columns(1).ColumnWidth = 70: ws.Columns(2).ColumnWidth = 26
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
    ' The formulas of the plan use LET, XLOOKUP, FILTER and TEXTSPLIT (Excel for Microsoft 365 or Excel 2024; Excel 2021 has no TEXTSPLIT, which the
    ' lists of FeatureTypes use): say so before anything is read or cleared.
    Cov "FLB35"
    If IsError(Application.Evaluate("LET(x,1,x)")) Or IsError(Application.Evaluate("COLUMNS(TEXTSPLIT(""a,b"","",""))")) Then
        Err.Raise vbObjectError + 708, "BuildFlowSheets", "This workbook needs Excel for Microsoft 365 or Excel 2024 (dynamic array formulas: LET, XLOOKUP, FILTER, TEXTSPLIT)."
    End If
End Sub

Private Sub LoadPlan(ByVal wb As Object, ByRef spec As Object, ByRef teams As Collection, ByRef iterations As Collection)
    ' Everything the build needs from the sheets, checked: the errors come before anything is changed.
    Dim nm As Variant
    RequireDynamicArrays
    For Each nm In Array("FlowFormulas", "ChartSpec", "ChecksSpec", "Config", "Iterations", "PIs", "Projects", "Workflows", _
                         "Transitions", "Dashboard", CHARTS_SHEET, ITERATION_CHARTS_SHEET, "Issues", "StatusIntervals", "IssueLinks", "PIIntervals")
        If Not SheetExists(wb, CStr(nm)) Then
            Err.Raise vbObjectError + 705, "BuildFlowSheets", "Sheet " & SheetName(CStr(nm)) & " is missing: rebuild the workbook with scripts\run.ps1 build."
        End If
    Next nm
    Set spec = ReadFormulas(wb)
    Set teams = TeamNames(wb.Worksheets(SheetName("Projects")), SettingOf(spec, "Dashboard", "ProgramTeam"))
    RequireColumns wb.Worksheets(SheetName("Iterations")), Array("Iteration-label", "Start", "End")
    RequireColumns wb.Worksheets(SheetName("PIs")), Array("PI", "Start", "End")
    Set iterations = IterationRows(wb.Worksheets(SheetName("Iterations")))
    If iterations.Count = 0 Then
        Cov "FLB15"
        Err.Raise vbObjectError + 703, "BuildFlowSheets", "Sheet Iterations has no iteration: fill Iteration-label, Start and End first."
    End If
    CheckPis wb.Worksheets(SheetName("PIs"))
    CheckModes wb.Worksheets(SheetName("Config"))
End Sub

Private Sub RequireColumns(ByVal ws As Object, ByVal columns As Variant)
    ' The headers that the build reads by name (ReadTable), as the Python path finds them (params_book.sheet_from_grid: trimmed, case kept). At the
    ' client (v0.19.0, AL 2026-11-26) a header that was not exactly Iteration-label left every label empty without a word: the keys of Metrics, of the
    ' check tables and of IterationFeatures were blank and 9 of the 20 controls of the Reconciliation were to review.
    ' The parts of the message are joined by ", ": " | " would cut it into fields of the Dashboard status (modText.StatusParts).
    Cov "FLB70"
    Dim have As Object, nm As Variant
    Set have = HeaderColumns(ws)
    For Each nm In columns
        If Not have.Exists(CStr(nm)) Then
            Cov "FLB71"
            Err.Raise vbObjectError + 713, "BuildFlowSheets", "Sheet " & ws.Name & ": no column " & nm & " in its first row (found: " & _
                      Join(have.Keys, ", ") & "). Type the header exactly, capitals and dashes included: " & Join(columns, ", ") & "."
        End If
    Next nm
End Sub

Private Sub CheckPis(ByVal wsPIs As Object)
    ' The dates of the rows of the sheet PIs that have a PI name (ReadTable reads no other row), as the Python path checks them (params/model.py): whole
    ' days from 2000 to 2099 (a date, or ISO text 2026-10-05), Start before End. Not here, unlike the Python path: a row without a name, a name twice. A PI date the formulas read and the VBA did not (a French text 05/10/2026), or the reverse (2026-10-05T00:00:00), made Feature
    ' Predictability not computable without a word (audit of the dates, 2026-11-23). A PI with no date at all is left out of the metrics, never a
    ' stop (AL, 2026-11-25: v0.18.2 stopped there with J2D-CFG-013 for PIs that Rovo saw in Jira without their dates): PiNames, modHealth.
    Cov "FLB64"
    Dim row As Variant, s As Variant, e As Variant, pi As String, bad As Boolean
    For Each row In ReadTable(wsPIs)
        pi = TrimWhite(CStr(row("PI") & ""))
        If PiWithoutDates(row("Start"), row("End")) Then
            Cov "FLB68"
        ElseIf Len(pi) > 0 Or Len(TrimWhite(CStr(row("Start") & ""))) > 0 Or Len(TrimWhite(CStr(row("End") & ""))) > 0 Then
            s = SerialOf(row("Start")): e = SerialOf(row("End"))
            If IsEmpty(s) Or IsEmpty(e) Then
                bad = True
            Else
                bad = (s < 36526 Or s >= 73051 Or e < 36526 Or e >= 73051 Or s <> Int(s) Or e <> Int(e) Or e <= s)
            End If
            If bad Then
                Cov "FLB65"
                Err.Raise vbObjectError + 710, "BuildFlowSheets", "Sheet PIs: Start and End of PI " & pi & " must be whole days from 2000 to 2099 " & _
                          "(a date, or 2026-10-05 as text), Start before End."
            End If
        End If
    Next row
End Sub

Private Sub CheckModes(ByVal wsConfig As Object)
    ' FlowLoadMode and PredictabilityMode are A or B, empty reads A (the Python path refuses anything else): "Mode A" made every Flow Load not
    ' computable without a word.
    Dim nm As Variant, v As String
    For Each nm In Array("FlowLoadMode", "PredictabilityMode")
        v = UCase$(ConfigValue(wsConfig, CStr(nm), "A"))
        If v <> "A" And v <> "B" Then
            Cov "FLB66"
            Err.Raise vbObjectError + 711, "BuildFlowSheets", "Config: " & CStr(nm) & " must be A or B, got """ & ConfigValue(wsConfig, CStr(nm)) & """."
        End If
    Next nm
End Sub

Public Sub ValidateBuildInputs(ByVal wb As Object)
    ' The inputs of the metrics (Iterations, PIs, the modes, an Excel with dynamic arrays): checked by 2. Compute the metrics before it changes
    ' anything, and by ExtractAndCompute before it reads Jira.
    Dim spec As Object, teams As Collection, iterations As Collection
    LoadPlan wb, spec, teams, iterations
End Sub

Public Sub BuildFlowSheetsIn(ByVal wb As Object)
    ' The tables and formulas, then the charts (the extraction in one go, the tests).
    Cov "FLB14"
    RunBuild wb, True, True
End Sub

Public Sub ComputeTablesIn(ByVal wb As Object)
    ' Button "2. Compute the metrics": the status rules made again (modExtract.RefreshRules) inside the build, once the old tables and the helper
    ' columns of the raw sheets are gone and the calculation is manual (written before, they recalculated the old tables several times), then the tables.
    RunBuild wb, True, False, True
End Sub

Public Sub BuildTablesIn(ByVal wb As Object)
    ' The tables and formulas from the raw sheets and the parameter sheets (button "2. Compute the metrics"); the charts are deleted: they read the
    ' tables that are made again.
    RunBuild wb, True, False
End Sub

Public Sub BuildChartsIn(ByVal wb As Object)
    ' The charts and their sheet ChartData from the tables (button "3. Draw the charts"); error 712 when the tables were not built.
    RunBuild wb, False, True
End Sub

Private Sub RunBuild(ByVal wb As Object, ByVal tables As Boolean, ByVal charts As Boolean, Optional ByVal refresh As Boolean = False)
    Dim spec As Object, teams As Collection, iterations As Collection, calcMode As Long, screen As Boolean
    Dim num As Long, desc As String
    LoadPlan wb, spec, teams, iterations
    If Not tables Then
        If Not TablesBuilt(wb) Then
            Cov "FLB67"
            Err.Raise vbObjectError + 712, "BuildFlowSheets", "The metrics are not computed: click 2. Compute the metrics first."
        End If
    End If
    calcMode = Application.Calculation
    screen = Application.ScreenUpdating
    On Error GoTo Failed                                              ' from here Failed gives both back
    Application.ScreenUpdating = False
    Application.Calculation = XL_CALC_MANUAL
    If tables Then BuildAll wb, spec, teams, iterations, refresh Else DeleteCharts wb
    If charts Then
        BuildCharts wb, spec, teams, FeatureTeamNames(wb, SettingOf(spec, "Dashboard", "ProgramTeam")), iterations, _
                    PiNames(wb, SettingOf(spec, "Dashboard", "NoPI")), StatusNames(wb)      ' its steps: "flow sheets: charts: ..." (one per chart)
    End If
    LogBegin "flow sheets: calculation"
    Application.Calculation = calcMode                   ' back to automatic: Excel calculates everything now
    If calcMode <> XL_CALC_AUTOMATIC Then Cov "FLB36": Application.CalculateFull      ' Excel left in manual mode by the user
    Application.ScreenUpdating = screen
    LogEnd "flow sheets: calculation"
    On Error Resume Next
    wb.Worksheets(SheetName("Dashboard")).Activate                  ' the sheets added by the build must not leave another one in front
    Exit Sub
Failed:
    num = Err.Number: desc = Err.Description
    Application.Calculation = calcMode
    Application.ScreenUpdating = screen
    Err.Raise num, "BuildFlowSheets", desc
End Sub

Private Function TablesBuilt(ByVal wb As Object) As Boolean
    ' The tables of 2. Compute the metrics were all made: the table Reconciliation is the last one (a build that stopped halfway left Metrics).
    Dim lo As Object
    If Not SheetExists(wb, "Reconciliation") Then Exit Function
    On Error Resume Next
    Set lo = wb.Worksheets(SheetName("Reconciliation")).ListObjects("Reconciliation")
    On Error GoTo 0
    TablesBuilt = Not lo Is Nothing
End Function

Private Sub BuildAll(ByVal wb As Object, ByVal spec As Object, ByVal teams As Collection, ByVal iterations As Collection, Optional ByVal refresh As Boolean = False)
    Cov "FLB16"
    Dim wsI As Object, wsS As Object, wsF As Object, wsD As Object, wsM As Object, wsC As Object, wsL As Object, wsR As Object
    Dim loI As Object, loS As Object, loF As Object, loD As Object, loM As Object, loC As Object, loL As Object, loR As Object
    Dim wsP As Object, loP As Object, pis As Collection, wsQ As Object, loQ As Object, wsT As Object, loT As Object
    Dim featureTeams As Collection, keys As Collection, statuses As Collection, wsU As Object, loU As Object, wsV As Object, loV As Object
    Dim wsFT As Object, loFT As Object, wsFE As Object, loFE As Object, wsFL As Object, loFL As Object, wsFD As Object, loFD As Object
    LogBegin "flow sheets: prepare"
    DeleteDerived wb
    Set wsI = wb.Worksheets(SheetName("Issues")): Set wsS = wb.Worksheets(SheetName("StatusIntervals"))
    PrepareRaw wsI, spec("Issues")
    PrepareRaw wsS, spec("StatusIntervals")
    Set wsL = wb.Worksheets(SheetName("IssueLinks"))
    PrepareRaw wsL, spec("IssueLinks")
    Set wsQ = wb.Worksheets(SheetName("PIIntervals"))
    PrepareRaw wsQ, spec("PIIntervals")
    Dim issues As Long, checks As Long, fills As Long
    issues = UsedRows(wsI)
    BuildControls wb, spec
    LogEnd "flow sheets: prepare"
    If refresh Then RefreshRules wb                                    ' its own log steps; the raw sheets are plain ranges here
    LogBegin "flow sheets: keys"
    Set wsF = AddSheet(wb, "ItemFlow"): Set wsD = AddSheet(wb, "DailyLoad"): Set wsM = AddSheet(wb, "Metrics")
    Set wsC = AddSheet(wb, "DataChecks"): Set wsR = AddSheet(wb, "FillRates"): Set wsT = AddSheet(wb, "IterationFeatures")
    Set wsP = AddSheet(wb, "PIMetrics"): Set wsU = AddSheet(wb, "PIStatusTime"): Set wsV = AddSheet(wb, "VelocityCheck")
    Set wsFT = AddSheet(wb, "FlowTimeCheck"): Set wsFE = AddSheet(wb, "FlowEfficiencyCheck"): Set wsFL = AddSheet(wb, "FlowLoadCheck")
    Set wsFD = AddSheet(wb, "FlowDistCheck")
    AddSheet wb, "MetricItems"
    AddSheet wb, "Reconciliation"
    Set pis = PiNames(wb, SettingOf(spec, "Dashboard", "NoPI"))
    Set statuses = StatusNames(wb)
    Set featureTeams = FeatureTeamNames(wb, SettingOf(spec, "Dashboard", "ProgramTeam"))
    Set keys = IterationFeatureKeys(wb, iterations, featureTeams)
    WriteDailyKeys wsD, iterations, teams
    WriteMetricKeys wsM, iterations, teams
    WriteMetricKeys wsV, iterations, teams                                      ' VelocityCheck has the keys of Metrics
    WriteMetricKeys wsFT, iterations, teams                                     ' FlowTimeCheck too
    WriteMetricKeys wsFE, iterations, teams                                     ' FlowEfficiencyCheck too
    WriteMetricKeys wsFL, iterations, teams                                     ' FlowLoadCheck too
    WriteMetricKeys wsFD, iterations, teams                                     ' FlowDistCheck too
    WritePiKeys wsP, pis
    WritePiStatusKeys wsU, pis, statuses
    WriteIterationFeatureKeys wsT, keys
    checks = WriteCheckKeys(wsC, wb.Worksheets(SheetName("ChecksSpec")))
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
    Set loQ = CreateTable(wsQ, "PIIntervals", spec("PIIntervals"), UsedRows(wsQ))
    Set loT = CreateTable(wsT, "IterationFeatures", spec("IterationFeatures"), keys.Count)
    Set loP = CreateTable(wsP, "PIMetrics", spec("PIMetrics"), pis.Count)
    Set loU = CreateTable(wsU, "PIStatusTime", spec("PIStatusTime"), pis.Count * statuses.Count)
    Set loV = CreateTable(wsV, "VelocityCheck", spec("VelocityCheck"), iterations.Count * teams.Count)
    Set loFT = CreateTable(wsFT, "FlowTimeCheck", spec("FlowTimeCheck"), iterations.Count * teams.Count)
    Set loFE = CreateTable(wsFE, "FlowEfficiencyCheck", spec("FlowEfficiencyCheck"), iterations.Count * teams.Count)
    Set loFL = CreateTable(wsFL, "FlowLoadCheck", spec("FlowLoadCheck"), iterations.Count * teams.Count)
    Set loFD = CreateTable(wsFD, "FlowDistCheck", spec("FlowDistCheck"), iterations.Count * teams.Count)
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
    LogBegin "flow sheets: formulas PIIntervals"
    FillFormulas loQ, spec("PIIntervals")
    LogEnd "flow sheets: formulas PIIntervals"
    LogBegin "flow sheets: formulas IterationFeatures"
    FillFormulas loT, spec("IterationFeatures")
    LogEnd "flow sheets: formulas IterationFeatures"
    LogBegin "flow sheets: formulas PIMetrics"
    FillFormulas loP, spec("PIMetrics")
    LogEnd "flow sheets: formulas PIMetrics"
    LogBegin "flow sheets: formulas PIStatusTime"
    FillFormulas loU, spec("PIStatusTime")
    LogEnd "flow sheets: formulas PIStatusTime"
    LogBegin "flow sheets: formulas VelocityCheck"
    FillFormulas loV, spec("VelocityCheck")
    LogEnd "flow sheets: formulas VelocityCheck"
    LogBegin "flow sheets: formulas FlowTimeCheck"
    FillFormulas loFT, spec("FlowTimeCheck")
    LogEnd "flow sheets: formulas FlowTimeCheck"
    LogBegin "flow sheets: formulas FlowEfficiencyCheck"
    FillFormulas loFE, spec("FlowEfficiencyCheck")
    LogEnd "flow sheets: formulas FlowEfficiencyCheck"
    LogBegin "flow sheets: formulas FlowLoadCheck"
    FillFormulas loFL, spec("FlowLoadCheck")
    LogEnd "flow sheets: formulas FlowLoadCheck"
    LogBegin "flow sheets: formulas FlowDistCheck"
    FillFormulas loFD, spec("FlowDistCheck")
    LogEnd "flow sheets: formulas FlowDistCheck"
    wsV.Range(wsV.Columns(2), wsV.Columns(3)).NumberFormat = "yyyy-mm-dd"
    wsFT.Range(wsFT.Columns(2), wsFT.Columns(3)).NumberFormat = "yyyy-mm-dd"
    wsFE.Range(wsFE.Columns(2), wsFE.Columns(3)).NumberFormat = "yyyy-mm-dd"
    wsFE.Range(wsFE.Columns(7), wsFE.Columns(8)).NumberFormat = "0.0"                            ' ActiveDays, WaitDays
    wsFE.Columns(10).NumberFormat = "0%"                                                       ' Efficiency
    wsFL.Range(wsFL.Columns(2), wsFL.Columns(3)).NumberFormat = "yyyy-mm-dd"
    wsFL.Columns(8).NumberFormat = "0.0"                                                       ' Average
    wsFD.Range(wsFD.Columns(2), wsFD.Columns(3)).NumberFormat = "yyyy-mm-dd"
    LogBegin "flow sheets: MetricItems"
    BuildMetricItems wb, spec
    LogEnd "flow sheets: MetricItems"
    LogBegin "flow sheets: Reconciliation"
    BuildReconciliation wb, spec
    LogEnd "flow sheets: Reconciliation"
    wsU.Range(wsU.Columns(3), wsU.Columns(4)).NumberFormat = "yyyy-mm-dd"
    wsU.Columns(7).NumberFormat = "0.0"                                          ' AverageDays
    wsU.Columns(8).NumberFormat = "0%"                                           ' Share
    wsD.Columns(1).NumberFormat = "yyyy-mm-dd"
    wsM.Range(wsM.Columns(2), wsM.Columns(3)).NumberFormat = "yyyy-mm-dd"
    wsP.Range(wsP.Columns(2), wsP.Columns(3)).NumberFormat = "yyyy-mm-dd"
    wsT.Range(wsT.Columns(2), wsT.Columns(3)).NumberFormat = "yyyy-mm-dd"
    wsT.Columns(6).NumberFormat = "yyyy-mm-dd"                                 ' Baseline
    wsQ.Range(wsQ.Columns(3), wsQ.Columns(4)).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    wsP.Columns(6).NumberFormat = "0%"
End Sub

Private Sub BuildControls(ByVal wb As Object, ByVal spec As Object)
    ' The as-of date and the Flow Load mode, read from Config: cells with names that the formulas use.
    Cov "FLB17"
    Dim ws As Object, i As Long
    Set ws = AddSheet(wb, "Controls")
    For i = 1 To spec("Controls").Count
        ws.Cells(i, 1).Value = spec("Controls")(i)(0)
        ws.Cells(i, 2).Formula2 = spec("Controls")(i)(1)
        wb.Names.Add spec("Controls")(i)(0), "=" & SheetName("Controls") & "!$B$" & i
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

Private Function FeatureTeamNames(ByVal wb As Object, ByVal programTeam As String) As Collection
    ' The ART as a whole, then the teams the Features name in their Team field (flow/iteration_features.feature_teams): the distinct
    ' TeamValue of the Features (type in FeatureTypes) of the sheet Issues, alphabetical (binary order, as Python sorted).
    Cov "FLB44"
    Dim out As New Collection, wsI As Object, last As Long, types As Variant, teamValues As Variant, features As Object, t As Variant
    Dim names As New Collection, seen As Object, r As Long, v As String, i As Long
    out.Add programTeam
    Set FeatureTeamNames = out
    Set wsI = wb.Worksheets(SheetName("Issues"))
    last = UsedRows(wsI)
    If last < 1 Then Exit Function
    Set features = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary")
    For Each t In Split(ConfigValue(wb.Worksheets(SheetName("Config")), "FeatureTypes"), ",")
        If Len(TrimWhite(CStr(t))) > 0 Then features(TrimWhite(CStr(t))) = True
    Next t
    types = IssueColumn(wsI, "Type", last)
    teamValues = IssueColumn(wsI, "TeamValue", last)
    For r = 1 To last
        v = TrimWhite(CStr(teamValues(r, 1) & ""))
        If Len(v) > 0 And v <> programTeam And features.Exists(CStr(types(r, 1) & "")) And Not seen.Exists(v) Then
            Cov "FLB45"
            seen(v) = True
            For i = 1 To names.Count
                If StrComp(v, CStr(names(i)), vbBinaryCompare) < 0 Then Exit For
            Next i
            If i > names.Count Then names.Add v Else names.Add v, , i
        End If
    Next r
    For i = 1 To names.Count
        out.Add names(i)
    Next i
End Function

Private Function IterationFeatureKeys(ByVal wb As Object, ByVal iterations As Collection, ByVal featureTeams As Collection) As Collection
    ' One row per iteration, PI that holds its start (the sheet PIs: Start <= iteration start < End, by Start then name: DatedPis; none = one row
    ' with an empty PI) and team of the Features: Array(label, start, end, team, PI). The counts are the formulas of the table IterationFeatures.
    Cov "FLB46"
    Dim out As New Collection, pisSheet As Collection, p As Variant, q As Variant, holders As Collection, t As Long
    Set pisSheet = DatedPis(wb)
    For Each p In iterations
        Set holders = New Collection
        For Each q In pisSheet
            If q(1) <= p(1) And p(1) < q(2) Then holders.Add q(0)
        Next q
        If holders.Count = 0 Then Cov "FLB47": holders.Add ""
        For Each q In holders
            For t = 1 To featureTeams.Count
                out.Add Array(CStr(p(0)), p(1), p(2), CStr(featureTeams(t)), CStr(q))
            Next t
        Next q
    Next p
    Set IterationFeatureKeys = out
End Function

Private Sub WriteIterationFeatureKeys(ByVal ws As Object, ByVal keys As Collection)
    ' Iteration-label | IterationStart | IterationEnd | Team | PI, as values (a formula typed in a table column fills the whole column).
    Cov "FLB48"
    Dim r As Long, data() As Variant
    ReDim data(1 To keys.Count, 1 To 5)
    For r = 1 To keys.Count
        data(r, 1) = AsText(CStr(keys(r)(0))): data(r, 2) = keys(r)(1): data(r, 3) = keys(r)(2): data(r, 4) = AsText(CStr(keys(r)(3)))
        If Len(CStr(keys(r)(4))) > 0 Then data(r, 5) = AsText(CStr(keys(r)(4)))
    Next r
    ws.Range(ws.Cells(2, 1), ws.Cells(keys.Count + 1, 5)).Value = data
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
    If col > lastCol Then Err.Raise vbObjectError + 709, "IssueColumn", "The sheet " & wsIssues.Name & " has no column " & header & ": run the extraction again."
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
    Set types = IssueTypes(wb.Worksheets(SheetName("Issues")))
    For Each pair In fields
        If Len(CStr(pair(1))) = 0 Then
            names.Add CStr(pair(0))
        ElseIf Len(ConfigValue(wb.Worksheets(SheetName("Config")), CStr(pair(1)))) > 0 Then
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

Private Sub BuildCharts(ByVal wb As Object, ByVal spec As Object, ByVal teams As Collection, ByVal featureTeams As Collection, _
                        ByVal iterations As Collection, ByVal pis As Collection, ByVal statuses As Collection)
    Cov "FLB21"
    Dim data As Object, dash As Object, row As Variant, index As Long, stride As Long
    Set dash = wb.Worksheets(SheetName(CHARTS_SHEET))               ' a sheet of their own: on the Dashboard they hid its cells
    Set data = AddSheet(wb, "ChartData")
    stride = IIf(pis.Count > iterations.Count, pis.Count, iterations.Count)     ' a block: one row per iteration or PI
    LogBegin "flow sheets: charts: selector | " & teams.Count & " teams, " & iterations.Count & " iterations, " & pis.Count & " PIs"
    WriteSelectorList wb, dash, data, spec, teams, featureTeams, statuses
    LogEnd "flow sheets: charts: selector"
    For Each row In ReadTable(wb.Worksheets(SheetName("ChartSpec")))
        If CStr(row("Sheet")) = ITERATION_CHARTS_SHEET Then
            Cov "FLB43"                                              ' the Features per iteration: their own sheet and the teams of the Features
            BuildChart wb.Worksheets(SheetName(ITERATION_CHARTS_SHEET)), data, row, 1 + index * (stride + 4), featureTeams, iterations, pis, statuses
        Else
            BuildChart dash, data, row, 1 + index * (stride + 4), teams, iterations, pis, statuses
        End If
        index = index + 1
    Next row
End Sub

Private Function PiNames(ByVal wb As Object, ByVal noPi As String) As Collection
    ' The PIs of the sheet PIs that have dates, by Start (AL, 2026-11-25: DatedPis), then those that Features carry but the sheet lacks or holds
    ' without dates (FeaturePis), in the natural order of their names; noPi alone when there is none, so that the chart still has a row (a gap).
    ' They are the keys of the table PIMetrics and the axis of the Feature Predictability chart (the table looks the dates up itself).
    ' Reference: flow/metrics.feature_predictability.
    Dim out As New Collection, seen As Object, q As Variant
    Cov "FLB37"
    Set seen = CreateObject("Scripting.Dictionary")
    For Each q In DatedPis(wb)
        out.Add CStr(q(0))
        seen(UCase$(CStr(q(0)))) = True                  ' MATCH and COUNTIFS of the table ignore the case
    Next q
    FeaturePis out, seen, wb
    If out.Count = 0 Then Cov "FLB38": out.Add noPi
    Set PiNames = out
End Function

Private Sub FeaturePis(ByVal out As Collection, ByVal seen As Object, ByVal wb As Object)
    ' Adds to out the PIs of the Features (type in FeatureTypes, case-sensitive like IsFeature) of the sheet Issues that the
    ' sheet PIs lacks or holds without dates (AL, 2026-11-12), the first spelling of each, in the natural order of the names (PiComesBefore):
    ' without dates they are not computable, but the RTE sees them. Any name is a PI (AL, 2026-11-25).
    Dim wsI As Object, last As Long, types As Variant, pis As Variant, features As Object, extra As New Collection
    Dim r As Long, j As Long, v As String, t As Variant
    Set wsI = wb.Worksheets(SheetName("Issues"))
    last = UsedRows(wsI)
    If last < 1 Then Exit Sub
    Set features = CreateObject("Scripting.Dictionary")              ' binary compare: case-sensitive, as FIND in IsFeature
    For Each t In Split(ConfigValue(wb.Worksheets(SheetName("Config")), "FeatureTypes"), ",")
        If Len(TrimWhite(CStr(t))) > 0 Then features(TrimWhite(CStr(t))) = True
    Next t
    types = IssueColumn(wsI, "Type", last)
    pis = IssueColumn(wsI, "PI", last)
    For r = 1 To last
        v = CStr(pis(r, 1) & "")
        If Len(v) > 0 Then
            If features.Exists(CStr(types(r, 1) & "")) And Not seen.Exists(UCase$(v)) Then
                Cov "FLB41"
                seen(UCase$(v)) = True
                For j = 1 To extra.Count
                    If PiComesBefore(v, 0, CStr(extra(j)), 0) Then Exit For
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

Private Function DatedPis(ByVal wb As Object) As Collection
    ' The PIs of the sheet PIs that have dates, Array(name, start, end), by Start, then in the natural order of the names (AL, 2026-11-25: the
    ' order of the rows does not count): the keys of PIMetrics (PiNames) and the PIs of an iteration (IterationFeatureKeys). Python:
    ' flow/model.ordered_pis.
    Dim out As New Collection, q As Variant, s As Variant, e As Variant, nm As String, i As Long
    For Each q In ReadTable(wb.Worksheets(SheetName("PIs")))
        s = SerialOf(q("Start")): e = SerialOf(q("End")): nm = Trim$(CStr(q("PI") & ""))
        If Len(nm) > 0 And Not IsEmpty(s) And Not IsEmpty(e) Then
            For i = 1 To out.Count
                If PiComesBefore(nm, CDbl(s), CStr(out(i)(0)), CDbl(out(i)(1))) Then Exit For
            Next i
            If i > out.Count Then
                out.Add Array(nm, s, e)
            Else
                Cov "FLB69"                                          ' before a later PI already read
                out.Add Array(nm, s, e), , i
            End If
        End If
    Next q
    Set DatedPis = out
End Function

Public Function PiComesBefore(ByVal nameA As String, ByVal startA As Double, ByVal nameB As String, ByVal startB As Double) As Boolean
    ' The order of the PIs (AL, 2026-11-25): by Start (0 for a PI without dates), then by PiSortKey, then by the name itself.
    Dim c As Long
    If startA <> startB Then PiComesBefore = (startA < startB): Exit Function
    c = StrComp(PiSortKey(nameA), PiSortKey(nameB), vbBinaryCompare)
    If c = 0 Then c = StrComp(nameA, nameB, vbBinaryCompare)
    PiComesBefore = (c < 0)
End Function

Public Function PiSortKey(ByVal name As String) As String
    ' The natural order of PI names: a run of digits compares as a number (PI#3-2024 before PI#10-2026), a to z as A to Z, compared as binary
    ' text. The same text as flow/model.pi_sort_key: a run of digits = character 1, its length on 4 digits, the digits without leading zeros.
    Dim i As Long, ch As String, digits As String, out As String, code As Long
    For i = 1 To Len(name)
        ch = Mid$(name, i, 1)
        code = AscW(ch)
        If code >= 48 And code <= 57 Then
            digits = digits & ch
        Else
            If Len(digits) > 0 Then out = out & DigitRun(digits): digits = ""
            If code >= 97 And code <= 122 Then ch = ChrW(code - 32)
            out = out & ch
        End If
    Next i
    If Len(digits) > 0 Then out = out & DigitRun(digits)
    PiSortKey = out
End Function

Private Function DigitRun(ByVal digits As String) As String
    Do While Left$(digits, 1) = "0"
        digits = Mid$(digits, 2)
    Loop
    DigitRun = Chr$(1) & Format$(Len(digits), "0000") & digits
End Function

Private Sub WriteSelectorList(ByVal wb As Object, ByVal dash As Object, ByVal data As Object, ByVal spec As Object, _
                              ByVal teams As Collection, ByVal featureTeams As Collection, ByVal statuses As Collection)
    ' The team of the Distribution chart is picked from a list: All teams, then the teams.
    Cov "FLB22"
    Dim i As Long, cell As Object, allTeams As String, found As Boolean, listColumn As Long
    allTeams = SettingOf(spec, "Dashboard", "AllTeams")
    Set cell = dash.Range(SettingOf(spec, "Dashboard", "SelectorCell"))
    listColumn = SelectorColumn(teams.Count, featureTeams.Count, statuses.Count)      ' right of the widest chart block (one column per team, work type or status)
    data.Cells(1, listColumn).Value = "Teams for the selector"
    data.Cells(2, listColumn).Value = AsText(allTeams)
    found = (CStr(cell.Value) = allTeams)
    For i = 1 To teams.Count
        data.Cells(2 + i, listColumn).Value = AsText(CStr(teams(i)))
        If CStr(cell.Value) = CStr(teams(i)) Then found = True
    Next i
    cell.Validation.Delete
    cell.Validation.Add 3, 1, 1, "=" & SheetName("ChartData") & "!" & data.Range(data.Cells(2, listColumn), data.Cells(teams.Count + 2, listColumn)).Address   ' list, stop on error, between
    If Not found Then Cov "FLB23": cell.Value = allTeams
End Sub

Private Sub BuildChart(ByVal dash As Object, ByVal data As Object, ByVal row As Variant, ByVal top As Long, _
                       ByVal teams As Collection, ByVal iterations As Collection, ByVal pis As Collection, ByVal statuses As Collection)
    Cov "FLB24"
    Dim names() As String, count As Long, i As Long, j As Long, first As Long, last As Long, point As String
    Dim area As Object, shape As Object, ch As Object, s As Object, rng As Object
    Dim rowsCount As Long, firstDay As Double, p As Variant, band() As String, isDay As Boolean, axis() As Variant
    Dim isPi As Boolean, stage As String, num As Long, desc As String
    On Error GoTo Failed                                     ' only to say where: the error goes on, and the step stays open in the log
    stage = "plan"
    If CStr(row("SeriesNames")) = STATUSES_TOKEN Then
        Cov "FLB52"                                          ' one series per Active / Wait status of the Feature workflow (StatusNames)
        ReDim names(0 To statuses.Count - 1)
        For i = 1 To statuses.Count: names(i - 1) = statuses(i): Next i
    ElseIf Len(CStr(row("SeriesNames"))) > 0 Then
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
    LogBegin "flow sheets: charts: " & CStr(row("Name")) & " | " & rowsCount & " rows, " & count & " series"
    stage = "axis"
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
    stage = "point formulas"
    point = Replace(CStr(row("PointFormula")), "{iteration}", "$A" & first)
    If InStr(point, "{series}") > 0 Then
        Cov "FLB25"                                          ' one formula per series: the name is part of the text
        For j = 1 To count
            data.Range(data.Cells(first, 1 + j), data.Cells(last, 1 + j)).Formula2 = Replace(point, "{series}", names(j - 1))
        Next j
    Else
        data.Range(data.Cells(first, 2), data.Cells(last, 1 + count)).Formula2 = Replace(point, "{name}", "B$" & (top + 1))
    End If
    stage = "chart object"
    Set area = dash.Range(dash.Cells(CLng(row("AnchorRow")), CLng(row("AnchorColumn"))), _
                          dash.Cells(CLng(row("AnchorRow")) + CLng(row("Rows")) - 1, CLng(row("AnchorColumn")) + CLng(row("Columns")) - 1))
    Set shape = dash.Shapes.AddChart2(-1, CLng(row("ChartType")), area.Left, area.Top, area.Width, area.Height, True)
    shape.Name = CStr(row("Name"))
    Set ch = shape.Chart
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop
    stage = "series"
    For j = 1 To count
        Set s = ch.SeriesCollection.NewSeries
        s.Name = "=" & SheetName("ChartData") & "!" & data.Cells(top + 1, 1 + j).Address
        s.Values = data.Range(data.Cells(first, 1 + j), data.Cells(last, 1 + j))
        s.XValues = data.Range(data.Cells(first, 1), data.Cells(last, 1))
    Next j
    If CStr(row("Labels") & "") = "yes" Then
        Cov "FLB49"                                          ' the value above each bar of the first series (the ART): a 0 has no bar to see
        stage = "labels"
        ch.SeriesCollection(1).HasDataLabels = True
        ch.SeriesCollection(1).DataLabels.Position = 2       ' xlLabelPositionOutsideEnd
    End If
    If Len(CStr(row("Band"))) > 0 Then
        Cov "FLB34"                                          ' a target band: two constant series, dashed, after the data
        stage = "band"
        band = Split(CStr(row("Band")), ";")
        For j = 0 To 1
            data.Cells(top + 1, 2 + count + j).Value = IIf(j = 0, "Target band low", "Target band high")
            data.Range(data.Cells(first, 2 + count + j), data.Cells(last, 2 + count + j)).Value = Val(band(j))
            Set s = ch.SeriesCollection.NewSeries
            s.Name = "=" & SheetName("ChartData") & "!" & data.Cells(top + 1, 2 + count + j).Address
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
    stage = "title and legend"
    ch.HasTitle = True
    ch.ChartTitle.Text = CStr(row("Title"))
    ch.HasLegend = True
    ch.Legend.Position = -4107                               ' xlLegendPositionBottom
    ch.DisplayBlanksAs = 1                                   ' xlNotPlotted: a figure that is not computable is a gap, never 0
    If Len(CStr(row("AxisFormat"))) > 0 Then Cov "FLB26": ch.Axes(2).TickLabels.NumberFormat = CStr(row("AxisFormat"))
    stage = "note"
    Set rng = data.Range(data.Cells(first, 2), data.Cells(last, 1 + count))
    dash.Cells(CLng(row("AnchorRow")) + CLng(row("Rows")), CLng(row("AnchorColumn"))).Formula2 = _
        Replace(CStr(row("NoteFormula")), "{range}", SheetName("ChartData") & "!" & rng.Address)
    LogEnd "flow sheets: charts: " & CStr(row("Name"))
    Exit Sub
Failed:
    ' The error goes on as it came; the log gets where it stopped, and numbers only (no name, no value of Jira).
    num = Err.Number: desc = Err.Description
    LogLine "WARN", "chart " & CStr(row("Name")) & " stopped at: " & stage & " | " & rowsCount & " rows, " & count & " series, " & _
            teams.Count & " teams, " & iterations.Count & " iterations, " & pis.Count & " PIs"
    Err.Raise num, "BuildChart", desc
End Sub
