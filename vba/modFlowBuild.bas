Attribute VB_Name = "modFlowBuild"
Option Explicit
' Builds the flow sheets (ItemFlow, DailyLoad, Metrics, Controls) and the dashboard (ChartData, the charts) after the
' extraction. Nothing is computed here: Excel formulas do it. The formulas and the description of the charts are NOT in
' this code: they are rows of the sheets FlowFormulas (Table | Column | Formula) and ChartSpec, written by the Python
' generator from the tested formula text. An empty Formula marks a column of values (raw data or keys written here).
' The sizes follow the data: Issues, StatusIntervals, ItemFlow = one row per extracted row, DailyLoad = one row per day
' and team, Metrics = one row per period and team. Every call rebuilds everything, so no row of an older run stays.

Private Const XL_SRC_RANGE As Long = 1
Private Const XL_YES As Long = 1
Private Const XL_CALC_MANUAL As Long = -4135

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
    If SheetExists(wb, "Dashboard") Then
        For i = wb.Worksheets("Dashboard").ChartObjects.Count To 1 Step -1
            wb.Worksheets("Dashboard").ChartObjects(i).Delete
        Next i
    End If
    If SheetExists(wb, "Dashboard") Then wb.Worksheets("Dashboard").Range("A4:R400").ClearContents      ' the notes under the charts
    alerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error GoTo Restore
    For Each nm In Array("ChartData", "ItemFlow", "DailyLoad", "Metrics", "DataChecks", "FillRates", "Controls")
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

Private Function PeriodRows(ByVal wsPeriods As Object) As Collection
    ' Period | Start | End -> Collection of Array(name, start serial, end serial)
    Cov "FLB09"
    Dim out As New Collection, row As Variant, s As Variant, e As Variant
    For Each row In ReadTable(wsPeriods)
        s = SerialOf(row("Start")): e = SerialOf(row("End"))
        If IsEmpty(s) Or IsEmpty(e) Then
            Cov "FLB28"
            Err.Raise vbObjectError + 701, "BuildFlowSheets", "Sheet Periods: Start and End of period " & row("Period") & " must be dates."
        End If
        If e <= s Then
            Cov "FLB27"
            Err.Raise vbObjectError + 704, "BuildFlowSheets", "Sheet Periods: period " & row("Period") & " ends before it starts."
        End If
        out.Add Array(CStr(row("Period")), s, e)
    Next row
    Set PeriodRows = out
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
    Dim i As Long, last As Long
    For i = 1 To cols.Count
        ws.Cells(1, i).Value = cols(i)(0)
    Next i
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

Public Sub BuildFlowSheetsIn(ByVal wb As Object)
    Cov "FLB14"
    Dim spec As Object, teams As Collection, periods As Collection, calcMode As Long
    Dim num As Long, desc As String, nm As Variant
    For Each nm In Array("FlowFormulas", "ChartSpec", "ChecksSpec", "Config", "Periods", "PIs", "Projects", "Workflows", _
                         "Transitions", "Dashboard", "Issues", "StatusIntervals", "IssueLinks")
        If Not SheetExists(wb, CStr(nm)) Then
            Err.Raise vbObjectError + 705, "BuildFlowSheets", "Sheet " & nm & " is missing: rebuild the workbook with scripts\run.ps1 build."
        End If
    Next nm
    Set spec = ReadFormulas(wb)
    Set teams = TeamNames(wb.Worksheets("Projects"), SettingOf(spec, "Dashboard", "ProgramTeam"))
    Set periods = PeriodRows(wb.Worksheets("Periods"))
    If periods.Count = 0 Then
        Cov "FLB15"
        Err.Raise vbObjectError + 703, "BuildFlowSheets", "Sheet Periods has no period: fill Period, Start and End first."
    End If
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = XL_CALC_MANUAL
    On Error GoTo Failed
    BuildAll wb, spec, teams, periods
    Application.Calculation = calcMode
    Application.CalculateFull
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    num = Err.Number: desc = Err.Description
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Err.Raise num, "BuildFlowSheets", desc
End Sub

Private Sub BuildAll(ByVal wb As Object, ByVal spec As Object, ByVal teams As Collection, ByVal periods As Collection)
    Cov "FLB16"
    Dim wsI As Object, wsS As Object, wsF As Object, wsD As Object, wsM As Object, wsC As Object, wsL As Object, wsR As Object
    Dim loI As Object, loS As Object, loF As Object, loD As Object, loM As Object, loC As Object, loL As Object, loR As Object
    DeleteDerived wb
    Set wsI = wb.Worksheets("Issues"): Set wsS = wb.Worksheets("StatusIntervals")
    PrepareRaw wsI, spec("Issues")
    PrepareRaw wsS, spec("StatusIntervals")
    Set wsL = wb.Worksheets("IssueLinks")
    PrepareRaw wsL, spec("IssueLinks")
    Dim issues As Long, checks As Long, fills As Long
    issues = UsedRows(wsI)
    BuildControls wb, spec
    Set wsF = AddSheet(wb, "ItemFlow"): Set wsD = AddSheet(wb, "DailyLoad"): Set wsM = AddSheet(wb, "Metrics")
    Set wsC = AddSheet(wb, "DataChecks"): Set wsR = AddSheet(wb, "FillRates")
    WriteDailyKeys wsD, periods, teams
    WriteMetricKeys wsM, periods, teams
    checks = WriteCheckKeys(wsC, wb.Worksheets("ChecksSpec"))
    fills = WriteFillKeys(wsR, spec("FillFields"), wb, teams)
    ' every table exists before any formula is written: the formulas refer to each other's columns
    Set loI = CreateTable(wsI, "Issues", spec("Issues"), issues)
    Set loS = CreateTable(wsS, "StatusIntervals", spec("StatusIntervals"), UsedRows(wsS))
    Set loF = CreateTable(wsF, "ItemFlow", spec("ItemFlow"), issues)
    Set loD = CreateTable(wsD, "DailyLoad", spec("DailyLoad"), DaysOf(periods) * teams.Count)
    Set loM = CreateTable(wsM, "Metrics", spec("Metrics"), periods.Count * teams.Count)
    Set loC = CreateTable(wsC, "DataChecks", spec("DataChecks"), checks)
    Set loL = CreateTable(wsL, "IssueLinks", spec("IssueLinks"), UsedRows(wsL))
    Set loR = CreateTable(wsR, "FillRates", spec("FillRates"), fills)
    FillFormulas loS, spec("StatusIntervals")
    FillFormulas loF, spec("ItemFlow")
    FillFormulas loL, spec("IssueLinks")
    FillFormulas loR, spec("FillRates")
    FillFormulas loC, spec("DataChecks")
    FillFormulas loD, spec("DailyLoad")
    FillFormulas loM, spec("Metrics")
    wsD.Columns(1).NumberFormat = "yyyy-mm-dd"
    wsM.Range(wsM.Columns(2), wsM.Columns(3)).NumberFormat = "yyyy-mm-dd"
    BuildCharts wb, spec, teams, periods
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

Private Function DaysOf(ByVal periods As Collection) As Long
    ' Whole days from the first period start to the last period end (end excluded).
    Cov "FLB18"
    Dim p As Variant, first As Double, last As Double
    first = Int(periods(1)(1)): last = Int(periods(1)(2))
    For Each p In periods
        If Int(p(1)) < first Then first = Int(p(1))
        If Int(p(2)) > last Then last = Int(p(2))
    Next p
    DaysOf = CLng(last - first)
End Function

Private Sub WriteDailyKeys(ByVal ws As Object, ByVal periods As Collection, ByVal teams As Collection)
    ' Date | Team, one row per day and team. Values, not formulas: a formula typed in a table column fills the whole column.
    Cov "FLB19"
    Dim p As Variant, first As Double, n As Long, d As Long, t As Long, r As Long, data() As Variant
    first = Int(periods(1)(1))
    For Each p In periods
        If Int(p(1)) < first Then first = Int(p(1))
    Next p
    n = DaysOf(periods) * teams.Count
    If n = 0 Then Exit Sub
    ReDim data(1 To n, 1 To 2)
    For d = 0 To DaysOf(periods) - 1
        For t = 1 To teams.Count
            r = r + 1
            data(r, 1) = first + d
            data(r, 2) = AsText(CStr(teams(t)))
        Next t
    Next d
    ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, 2)).Value = data
End Sub

Private Sub WriteMetricKeys(ByVal ws As Object, ByVal periods As Collection, ByVal teams As Collection)
    ' Period | PeriodStart | PeriodEnd | Team, one row per period and team.
    Cov "FLB20"
    Dim p As Variant, t As Long, r As Long, data() As Variant
    ReDim data(1 To periods.Count * teams.Count, 1 To 4)
    For Each p In periods
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
    Dim out As New Collection, seen As Object, col As Long, last As Long, v As String, r As Long
    Set seen = CreateObject("Scripting.Dictionary")
    Set IssueTypes = out
    last = UsedRows(wsIssues)
    If last < 1 Then Exit Function
    For col = 1 To 20
        If CStr(wsIssues.Cells(1, col).Value) = "Type" Then Exit For
    Next col
    For r = 1 To last                                ' cell by cell: a one-row range does not come back as an array
        v = CStr(wsIssues.Cells(r + 1, col).Value & "")
        If Len(v) > 0 Then
            If Not seen.Exists(v) Then seen(v) = True: out.Add v
        End If
    Next r
End Function

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

Private Sub BuildCharts(ByVal wb As Object, ByVal spec As Object, ByVal teams As Collection, ByVal periods As Collection)
    Cov "FLB21"
    Dim data As Object, dash As Object, row As Variant, index As Long
    Set dash = wb.Worksheets("Dashboard")
    Set data = AddSheet(wb, "ChartData")
    WriteSelectorList wb, dash, data, spec, teams
    For Each row In ReadTable(wb.Worksheets("ChartSpec"))
        BuildChart dash, data, row, 1 + index * (periods.Count + 4), teams, periods
        index = index + 1
    Next row
End Sub

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
                       ByVal teams As Collection, ByVal periods As Collection)
    Cov "FLB24"
    Dim names() As String, count As Long, i As Long, j As Long, first As Long, last As Long, point As String
    Dim area As Object, shape As Object, ch As Object, s As Object, rng As Object
    If Len(CStr(row("SeriesNames"))) > 0 Then
        names = Split(CStr(row("SeriesNames")), ",")
    Else
        ReDim names(0 To teams.Count - 1)
        For i = 1 To teams.Count: names(i - 1) = teams(i): Next i
    End If
    count = UBound(names) + 1
    first = top + 2: last = first + periods.Count - 1
    data.Cells(top, 1).Value = CStr(row("Title"))
    data.Cells(top, 1).Font.Bold = True
    data.Cells(top + 1, 1).Value = "Period"
    For i = 1 To periods.Count
        data.Cells(first + i - 1, 1).Value = AsText(CStr(periods(i)(0)))
    Next i
    For j = 1 To count
        data.Cells(top + 1, 1 + j).Value = AsText(names(j - 1))
    Next j
    point = Replace(CStr(row("PointFormula")), "{period}", "$A" & first)
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
