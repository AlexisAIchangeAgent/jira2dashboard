Attribute VB_Name = "modActions"
Option Explicit
' The macros behind the buttons of the Dashboard, three steps (AL, 2026-11-22: correct and compute again without reading Jira):
'   1. Extract from Jira     (BtnExtract)          reads Jira into the raw sheets; the results of the previous computation are deleted
'   2. Compute the metrics   (BtnComputeMetrics)   the status rules again (modExtract.RefreshRules), then the tables and formulas; no call to Jira
'   3. Draw the charts       (BtnDrawCharts)       the charts from the tables; no call to Jira
' then VBA code for Rovo (BtnRovoDebug, last): the code of the workbook in one text file for Rovo (modRovoDebug); no call to Jira
' Trivial entry points (like modButtons): the scenarios of flow_excel.feature run them against the fake Jira.
' A button macro must be a Public Sub without arguments.
'
' The extraction works in slices: BtnExtract checks the account (GET myself), starts the extraction job (modExtract) and returns; one step of the job (a page of the
' search, a chunk of changelogs...) then runs per Application.OnTime call, so Excel stays usable, the bar of the Dashboard moves and the button
' Cancel works between two steps (and during a wait for Jira: modHttp reads gHttpCancel). Nothing of an earlier run is touched before the raw
' sheets are written (the last step of the extraction), so a cancel before that keeps the previous results. Compute and Draw take a few seconds
' each, in one go.
' ExtractAndCompute is the three steps in one go (no bar, no cancel): the tests and other macros use it.
' The workbook handler Workbook_BeforeClose (added by the build) calls AbortExtraction: a slice still scheduled when the
' workbook is closed would otherwise reopen it.

Private mBusy As Boolean
Private mCancel As Boolean
Private mExtracting As Boolean              ' True while Jira is read (extraction job or configuration): Cancel works
Private mNextSlice As Date
Private Const STATUS_ROWS As Long = 20            ' rows of fields under the status cell (dashboard_spec.STATUS_ROWS)
Private Const CHARTS_SHEET As String = "Charts"    ' the sheet of the charts (flow/dashboard_spec.CHARTS_SHEET)
Private Const EXTRACT_CAPTION As String = "1. Extract from Jira"        ' the captions of the buttons (flow/dashboard_spec)
Private Const COMPUTE_CAPTION As String = "2. Compute the metrics"
Private Const DRAW_CAPTION As String = "3. Draw the charts"
Private Const ROVO_DEBUG_CAPTION As String = "VBA code for Rovo"
' The mandatory parameters the metrics need (the extraction checks its own: modExtract.EXTRACT_PARAMETERS)
Private Const COMPUTE_PARAMETERS As String = "FlowLoadMode,PredictabilityMode,AsOf,FeatureTypes,FillRateThreshold"
Private mNextProc As String                ' the string given to OnTime: Schedule:=False needs the same one

Private Function SliceName() As String
    SliceName = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!ExtractSlice"     ' qualified: another open workbook must not answer the call
End Function

Public Sub ExtractAndCompute(Optional ByVal fromButton As Boolean = False)
    ' The optional argument keeps this Sub out of the Alt+F8 list: only the buttons (with their error handler) are for users.
    ' Everything comes from the workbook: Config (connection, proxy, filter, as-of date), Projects, Iterations, Workflows.
    Dim num As Long, desc As String
    If mBusy Then Err.Raise vbObjectError + 611, "ExtractAndCompute", "An extraction is already running: wait for it or cancel it."
    On Error GoTo Failed
    mBusy = True                                   ' a click during a DoEvents must not start a second job
    LogRunStart "macro ExtractAndCompute", LogLevelSetting()
    LogParameters ThisWorkbook                     ' the inputs of the run (no value of the client: modLog.ParameterShown)
    CheckComputeInputs                             ' before Jira: a bad Iterations sheet must not be found after the reading
    Application.StatusBar = "Reading Jira..."
    ExtractFlowData                                ' empties the old results only once Jira has answered
    Application.StatusBar = "Computing the metrics..."
    ComputeMetricsCore
    Application.StatusBar = "Drawing the charts..."
    BuildChartsIn ThisWorkbook
    Application.StatusBar = False
    mBusy = False
    LogRunEnd "ok"
    Exit Sub
Failed:
    num = Err.Number: desc = Err.Description
    LogError "ExtractAndCompute", num, desc
    LogRunEnd "stopped"
    ExtractJobAbort                                ' releases the connection (Authorization header) and the parsed pages
    mBusy = False
    Application.StatusBar = False
    Err.Raise num, "ExtractAndCompute", desc
End Sub

Public Sub BtnExtract()
    ' 1. Extract from Jira: the job in slices (ExtractSlice). Its status says what was read and that 2. Compute the metrics comes next.
    Dim num As Long, desc As String
    If mBusy Then Beep: Exit Sub
    On Error GoTo Failed
    mBusy = True: mCancel = False: mExtracting = True: gHttpCancel = False
    Randomize                                                       ' the retry jitter must differ from one run to the next
    LogRunStart "button " & EXTRACT_CAPTION, LogLevelSetting()
    LogParameters ThisWorkbook                                      ' the inputs of the run (no value of the client)
    RequireEditable                                                 ' a locked workbook (Protected View, read-only) stops here, before Jira
    ShowProgress 0, "Starting..."
    ExtractFlowDataStart
    ScheduleSlice
    Exit Sub
Failed:
    num = Err.Number: desc = Err.Description                        ' first: the calls below may clear Err
    ExtractJobAbort
    If num = vbObjectError + 610 Then                                ' the Cancel button, pressed during the check of the account
        LogRunEnd "cancelled during a wait for Jira"
        EndRun "Cancelled | Results: the results of the previous run are kept", 0
        Exit Sub
    End If
    LogError "BtnExtract", num, desc
    LogRunEnd "stopped"
    EndRunStopped num, desc
End Sub

Public Sub BtnComputeMetrics()
    ' 2. Compute the metrics: the status rules again (the sheet Workflows, the roles of Projects, AsOf), then the tables and formulas from the raw
    ' sheets. No call to Jira, one go. The charts are deleted: they read the tables made again (3. Draw the charts).
    Dim num As Long, desc As String, health As String
    If mBusy Then Beep: Exit Sub
    On Error GoTo Failed
    mBusy = True: mCancel = False: mExtracting = False
    LogRunStart "button " & COMPUTE_CAPTION, LogLevelSetting()
    LogParameters ThisWorkbook                                      ' the inputs of the run (no value of the client)
    RequireEditable
    ShowProgress 0, "Computing the metrics..."
    ComputeMetricsCore
    health = HealthFields(ThisWorkbook)                             ' before the end of the run: its step and its lines belong to it
    LogRunEnd "ok"
    EndRun "Computed | Figures: the sheet " & SheetName("Metrics") & " | Data checks: the sheet " & SheetName("DataChecks") & _
           " | Next: click " & DRAW_CAPTION & health, 1
    Exit Sub
Failed:
    num = Err.Number: desc = Err.Description
    LogError "BtnComputeMetrics", num, desc
    LogRunEnd "stopped"
    EndRunStopped num, desc
End Sub

Public Sub BtnDrawCharts()
    ' 3. Draw the charts from the tables of 2. Compute the metrics (error 712 without them). No call to Jira, one go.
    Dim num As Long, desc As String, health As String
    If mBusy Then Beep: Exit Sub
    On Error GoTo Failed
    mBusy = True: mCancel = False: mExtracting = False
    LogRunStart "button " & DRAW_CAPTION, LogLevelSetting()
    LogParameters ThisWorkbook
    RequireEditable
    ShowProgress 0, "Drawing the charts..."
    BuildChartsIn ThisWorkbook
    health = HealthFields(ThisWorkbook)                             ' what to look at, again: the last status of the three steps says it
    LogRunEnd "ok"
    EndRun "Drawn | Charts: the six flow metrics in the sheet " & SheetName(CHARTS_SHEET) & " | Figures: the sheet " & SheetName("Metrics") & _
           " | Data checks: the sheet " & SheetName("DataChecks") & health, 1
    Exit Sub
Failed:
    num = Err.Number: desc = Err.Description
    LogError "BtnDrawCharts", num, desc
    LogRunEnd "stopped"
    EndRunStopped num, desc
End Sub

Private Sub ComputeMetricsCore()
    ' The inputs of the metrics checked first (nothing changes before), then the rules and the tables (modFlowBuild.ComputeTablesIn: RefreshRules inside).
    CheckComputeInputs
    ComputeTablesIn ThisWorkbook
End Sub

Private Sub CheckComputeInputs()
    ' The mandatory values of the metrics, AsOf readable, Iterations, PIs, the modes: errors before any change.
    Dim asOf As Date
    RequireMandatory ThisWorkbook.Worksheets(SheetName("Config")), COMPUTE_PARAMETERS
    asOf = AsOfFromConfig(ThisWorkbook.Worksheets(SheetName("Config")))
    ValidateBuildInputs ThisWorkbook
End Sub

Private Function RawRows(ByVal name As String) As Long
    ' The rows of a raw sheet (its first column is the key); 0 when it cannot be read.
    Dim ws As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SheetName(name))
    RawRows = ws.Cells(ws.Rows.Count, 1).End(-4162).Row - 1
    If RawRows < 0 Then RawRows = 0
    Err.Clear
End Function

Public Sub BtnCancelExtraction()
    ' Takes effect before the next step of the extraction, or at once during a wait for Jira; too late once the raw sheets
    ' are written.
    If Not mBusy Then Exit Sub
    If mExtracting Then
        mCancel = True
        gHttpCancel = True
        LogLine "RUN", "cancel asked"
        ShowProgress -1, "Cancelling..."
    Else
        LogLine "RUN", "cancel asked: nothing to cancel (this step runs in one go)"
        ShowProgress -1, "Nothing to cancel: this step runs in one go."
    End If
End Sub

Public Sub BtnReadJiraConfiguration()
    RunJiraConfiguration False
End Sub

Public Sub BtnProposeValues()
    RunJiraConfiguration True
End Sub

Private Sub RunJiraConfiguration(ByVal propose As Boolean)
    ' Reads the configuration of Jira (read-only) into the proposal sheets JiraIssueTypes, JiraStatuses, JiraWorkflows,
    ' JiraTransitions and JiraFields (modJiraConfig). One go, about two calls per project: no slices. Cancel works during
    ' a wait for Jira; the proposal sheets are written only at the end, so a cancel or an error leaves them as they were.
    Dim num As Long, desc As String, summary As String, withWorkflows As Boolean, cfg As Object, keys As Collection
    If mBusy Then Beep: Exit Sub
    On Error GoTo Failed
    mBusy = True: mCancel = False: mExtracting = True: gHttpCancel = False
    Randomize                                                       ' the retry jitter must differ from one run to the next
    LogRunStart IIf(propose, "button Propose values from Jira", "button Read the Jira configuration"), LogLevelSetting()
    LogParameters ThisWorkbook                                      ' the inputs of the run (no value of the client)
    RequireEditable                                                 ' a locked workbook (Protected View, read-only) stops here, before Jira
    ShowProgress 0, "Reading the Jira configuration..."
    RequireMandatory ThisWorkbook.Worksheets(SheetName("Config")), "BaseUrl,Email,ApiToken"      ' only the connection is needed here
    withWorkflows = IsYes(ConfigValue(ThisWorkbook.Worksheets(SheetName("Config")), "IsUserProjectAdmin", "True"))
    Set cfg = JiraConfigFromWorkbook()
    Set keys = ProjectKeys(ThisWorkbook.Worksheets(SheetName("Projects")))
    If keys.Count > 0 Then RequireLogin cfg                         ' a wrong account stops here (Jira Cloud would answer empty lists); an
                                                                    ' empty Projects stops in ReadJiraConfiguration (620), with no call
    If propose Then
        summary = ProposeValues(cfg, keys, ThisWorkbook, withWorkflows)
    Else
        summary = ReadJiraConfiguration(cfg, keys, ThisWorkbook, withWorkflows)
    End If
    LogRunEnd "ok"
    mExtracting = False
    EndRun summary, 1
    Exit Sub
Failed:
    num = Err.Number: desc = Err.Description                        ' first: the calls below may clear Err
    mExtracting = False
    If num = vbObjectError + 610 Then                                ' the Cancel button, pressed during a wait for Jira
        LogRunEnd "cancelled during a wait for Jira"
        EndRun "Cancelled | Proposal sheets: kept as they were", 0
        Exit Sub
    End If
    LogError IIf(propose, "BtnProposeValues", "BtnReadJiraConfiguration"), num, desc
    LogRunEnd "stopped"
    EndRunStopped num, desc
End Sub

Public Sub BtnRovoPrompt()
    ' Writes the prompt for Rovo into the sheet RovoPrompt (modParams.WriteRovoPrompt) and selects its lines: Ctrl+C, then paste
    ' into Rovo. The workbook never sends it anywhere. Not while an extraction runs (the sheet would change under it).
    Dim num As Long, desc As String, n As Long, ws As Object
    If mBusy Then Beep: Exit Sub
    On Error GoTo Failed
    LogRunStart "button Prompt for Rovo", LogLevelSetting()            ' a run of the log, as the other buttons
    LogParameters ThisWorkbook
    RequireEditable
    LogBegin "prompt for Rovo"
    n = WriteRovoPrompt(ThisWorkbook)
    LogEnd "prompt for Rovo: " & n & " lines"
    Set ws = ThisWorkbook.Worksheets(SheetName(SHEET_ROVO))
    ws.Activate
    ws.Range(ws.Cells(1, 1), ws.Cells(n, 1)).Select
    ShowProgress -1, "Rovo prompt | Lines: " & n & ", selected in the sheet " & SheetName(SHEET_ROVO) & _
                     " | Next: Ctrl+C, paste into Rovo, save its CSV answer, then click Load settings CSV and choose the file"
    LogRunEnd "ok"
    Exit Sub
Failed:
    num = Err.Number: desc = Err.Description
    LogError "BtnRovoPrompt", num, desc
    LogRunEnd "stopped"
    EndRunStopped num, desc
End Sub

Public Sub BtnRovoDebug()
    ' VBA code for Rovo (AL, 2026-11-27: Rovo reads an xlsx, not the macros of an xlsm): writes vba-code.txt into the folder rovo-debug next to the
    ' workbook (modRovoDebug.WriteVbaCode): the code of the delivered modules and the formulas of the tables, read from the sheets VbaCode and
    ' FlowFormulas (never from the VBA project). The workbook sends nothing: the user uploads the file into Rovo Chat. The log gets counts only.
    Dim num As Long, desc As String, folder As String, parts() As String
    If mBusy Then Beep: Exit Sub
    On Error GoTo Failed
    LogRunStart "button " & ROVO_DEBUG_CAPTION, LogLevelSetting()
    LogParameters ThisWorkbook
    RequireEditable
    folder = RovoDebugFolder(ThisWorkbook.Path, Environ$("LOCALAPPDATA"), Environ$("TEMP"))
    LogBegin "VBA code for Rovo"
    parts = Split(WriteVbaCode(ThisWorkbook, folder), "|")
    LogEnd "VBA code for Rovo: " & parts(0) & " modules, " & parts(1) & " lines, " & parts(3) & " formulas, " & parts(2) & " characters"
    ShowProgress -1, "VBA code for Rovo | File: " & VBA_CODE_FILE & ", " & parts(0) & " modules, " & parts(1) & " lines of code, " & parts(3) & _
                     " formulas, " & parts(2) & " characters | Folder: " & folder & " | Next: in Rovo Chat, + then Upload a file, then ask why a figure is what it is"
    LogRunEnd "ok"
    Exit Sub
Failed:
    num = Err.Number: desc = Err.Description
    LogError "BtnRovoDebug", num, desc
    LogRunEnd "stopped"
    EndRunStopped num, desc
End Sub

Public Sub BtnLoadRovoCsv()
    ' Asks for the CSV that Rovo answered (the prompt of the button Prompt for Rovo) in a dialog that opens in the folder of
    ' the workbook (the file may be anywhere), then loads its values into the parameter sheets (modParams.ImportRovoCsv). The
    ' user confirms first: the rows of the sheets in the file replace the current rows.
    Dim num As Long, desc As String, file As Variant, summary As String, folder As String
    If mBusy Then Beep: Exit Sub
    On Error GoTo Failed
    LogRunStart "button Load settings CSV", LogLevelSetting()         ' every click is a run of the log, a cancel too
    LogParameters ThisWorkbook                                       ' the values before the load
    RequireEditable
    folder = ThisWorkbook.Path
    If Len(folder) > 0 Then
        On Error Resume Next
        ChDrive folder: ChDir folder                                     ' the dialog opens in the folder of the workbook
        On Error GoTo Failed
    End If
    file = Application.GetOpenFilename("CSV files (*.csv),*.csv", , "The CSV answered by Rovo")
    If VarType(file) = vbBoolean Then
        ShowProgress -1, "Settings CSV | Result: no file chosen, nothing loaded"
        LogRunEnd "cancelled: no file chosen"
        Exit Sub
    End If
    If MsgBox("Load " & Mid$(file, InStrRev(file, "\") + 1) & " into the parameter sheets?" & vbLf & vbLf & _
              "The Value of each parameter of " & SheetName("Config") & " that the file gives replaces the current one. For every " & _
              "other sheet in the file (" & SheetName("Projects") & ", " & SheetName("Iterations") & ", " & SheetName("PIs") & ", " & _
              SheetName("Workflows") & ", " & SheetName("Transitions") & ") all the current rows are replaced by its rows. " & _
              "Save a copy of the workbook first if you want to keep the current values.", vbOKCancel + vbQuestion, _
              "Load settings CSV") <> vbOK Then
        ShowProgress -1, "Settings CSV | Result: cancelled, nothing loaded"
        LogRunEnd "cancelled: not confirmed"
        Exit Sub
    End If
    summary = ImportRovoCsvFile(CStr(file))
    ShowProgress -1, summary
    LogRunEnd "ok"
    Exit Sub
Failed:
    num = Err.Number: desc = Err.Description
    LogError "BtnLoadRovoCsv", num, desc
    LogRunEnd "stopped"
    EndRunStopped num, desc
End Sub

Public Function ImportRovoCsvFile(ByVal path As String) As String
    ' The work of BtnLoadRovoCsv without the dialogs (the tests and other macros use it): the step "settings CSV" of the run of
    ' the button. Called alone (no run in progress), it is a run of its own, as a button: its RUN line, the parameters before
    ' the load, the step and its time, RUN end. An error goes up to the caller. The summary is not logged: it quotes names
    ' read in the CSV.
    Dim summary As String, own As Boolean
    own = Not LogRunActive()
    If own Then
        LogRunStart "button Load settings CSV", LogLevelSetting()
        LogParameters ThisWorkbook                                   ' the values before the load
    End If
    LogBegin "settings CSV"
    summary = ImportRovoCsv(ThisWorkbook, path)
    LogEnd "settings CSV loaded"
    If own Then LogRunEnd "ok"
    ImportRovoCsvFile = summary
End Function

Public Sub ExtractSlice()
    ' Called by Application.OnTime: one step, then the next call is scheduled. Keep it short (UI thread).
    Dim num As Long, desc As String, health As String
    If Not mBusy Then Exit Sub
    On Error GoTo Failed
    If mExtracting Then
        If mCancel Then
            LogRunEnd "cancelled"
            ExtractJobAbort
            EndRun "Cancelled | Results: the results of the previous run are kept", 0
            Exit Sub
        End If
        If ExtractJobStep() Then
            mExtracting = False
            health = HealthFields(ThisWorkbook)                    ' the column types of the raw sheets (the tables come with 2. Compute)
            LogRunEnd "ok"
            EndRun "Extracted | Issues: " & RawRows("Issues") & " | Status intervals: " & RawRows("StatusIntervals") & " | Changelog rows: " & _
                   RawRows("Changelog") & " | Next: click " & COMPUTE_CAPTION & health, 1
            Exit Sub
        End If
        ShowProgress ExtractJobFraction(), ExtractJobLabel()
        If mBusy Then ScheduleSlice                ' AbortExtraction may have run during the DoEvents above
    End If
    Exit Sub
Failed:
    num = Err.Number: desc = Err.Description                        ' first: the calls below may clear Err
    If num = vbObjectError + 610 Then                                ' the Cancel button, pressed during a wait for Jira
        LogRunEnd "cancelled during a wait for Jira"
        ExtractJobAbort
        EndRun "Cancelled | Results: the results of the previous run are kept", 0
        Exit Sub
    End If
    LogError "ExtractSlice", num, desc
    LogRunEnd "stopped"
    ExtractJobAbort
    EndRunStopped num, desc
End Sub

Public Sub AbortExtraction()
    ' Forgets a pending slice (used when the workbook is closed or a test ends): an OnTime left behind would reopen the file.
    On Error Resume Next
    If mNextSlice <> 0 Then Application.OnTime mNextSlice, mNextProc, , False
    Err.Clear
    On Error GoTo 0
    mNextSlice = 0
    ExtractJobAbort
    mBusy = False: mCancel = False: mExtracting = False: gHttpCancel = False
    Application.StatusBar = False
End Sub

Public Function ProgressSnapshot() As String
    ' "fraction|status text" as the Dashboard shows them (one call for the tests that watch the bar from outside): the text is
    ' read back from the cells, the headline (the cell dashStatus) then each row under it as " | Name: value".
    Dim shapes As Object, cell As Object, text As String, i As Long, name As String, value As String
    On Error Resume Next
    Set shapes = ThisWorkbook.Worksheets(SheetName("Dashboard")).Shapes
    Set cell = ThisWorkbook.Names("dashStatus").RefersToRange
    text = CStr(cell.Value)
    For i = 1 To STATUS_ROWS
        name = CStr(cell.Offset(i, 0).Value): value = CStr(cell.Offset(i, 1).Value)
        If Len(name) = 0 And Len(value) = 0 Then Exit For
        text = text & " | " & IIf(Len(name) > 0, name & ": ", "") & value
    Next i
    ProgressSnapshot = Trim$(Str$(shapes("pbFill").Width / shapes("pbBack").Width)) & "|" & text
End Function

Private Function LogLevelSetting() As String
    ' Config LogLevel (Steps or Errors); Steps when the sheet or the row is missing.
    On Error Resume Next
    LogLevelSetting = "Steps"
    LogLevelSetting = ConfigValue(ThisWorkbook.Worksheets(SheetName("Config")), "LogLevel", "Steps")
    Err.Clear
End Function

Private Function OpenLogSetting() As Boolean
    ' Config OpenLogOnError (Yes or No); Yes when the sheet or the row is missing.
    On Error Resume Next
    OpenLogSetting = True
    OpenLogSetting = IsYes(ConfigValue(ThisWorkbook.Worksheets(SheetName("Config")), "OpenLogOnError", "Yes"))
    Err.Clear
End Function

Private Sub ScheduleSlice()
    mNextSlice = Now
    mNextProc = SliceName()
    Application.OnTime mNextSlice, mNextProc
End Sub

Private Sub EndRun(ByVal message As String, ByVal fraction As Double)
    mBusy = False: mCancel = False: mNextSlice = 0: gHttpCancel = False
    Application.StatusBar = False
    ShowProgress fraction, message
End Sub

Private Sub EndRunStopped(ByVal num As Long, ByVal desc As String)
    ' An error ends the run: the status gives the support code, the message and where the log is (one row each), and the log
    ' opens in Notepad (modLog.LogOpen).
    EndRun "Stopped | Code: " & ErrorCode(num, desc) & " | Message: " & desc & LogHint(), 0
    LogOpen OpenLogSetting()
End Sub

Private Function StatusValue(ByVal text As String) As Variant
    ' A value of a status row: a whole number as a number (Excel marks a number stored as text), any other text as text.
    If Len(text) > 0 And Len(text) <= 9 And text Like String$(Len(text), "#") Then StatusValue = CLng(text): Exit Function
    StatusValue = SafeCell(text)
End Function

Private Sub ShowProgress(ByVal fraction As Double, ByVal text As String)
    ' The bar (rectangles pbBack and pbFill) and the status of the Dashboard; a fraction < 0 keeps the bar. A text
    ' "Title | Name: value | ..." (modText.StatusParts) shows its title in the status cell (name dashStatus) and one row per
    ' field under it, the name in column A and the value in column B; the rows of the previous status are cleared first.
    Dim shapes As Object, cell As Object, parts As Collection, i As Long
    On Error Resume Next
    Set shapes = ThisWorkbook.Worksheets(SheetName("Dashboard")).Shapes
    If fraction >= 0 Then shapes("pbFill").Width = shapes("pbBack").Width * fraction + 0.1
    Set cell = ThisWorkbook.Names("dashStatus").RefersToRange
    Set parts = StatusParts(text)
    cell.Value = SafeCell(CStr(parts(1)(1)))
    ' The rows are touched only when there are fields to write or rows to clear: during a run each slice writes the headline only
    ' (more writes per slice made the known crash of the extraction come back, 1 run in 3).
    If parts.Count > 1 Or Len(CStr(cell.Offset(1, 0).Value)) > 0 Or Len(CStr(cell.Offset(1, 1).Value)) > 0 Then
        cell.Offset(1, 0).Resize(STATUS_ROWS, 2).ClearContents
    End If
    For i = 2 To parts.Count
        If i - 1 > STATUS_ROWS Then Exit For
        cell.Offset(i - 1, 0).Value = SafeCell(CStr(parts(i)(0)))
        cell.Offset(i - 1, 1).Value = StatusValue(CStr(parts(i)(1)))
    Next i
    Application.StatusBar = text
    If fraction >= 1 Or Not mBusy Then Application.StatusBar = False
    DoEvents
    Err.Clear
End Sub
