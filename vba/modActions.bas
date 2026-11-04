Attribute VB_Name = "modActions"
Option Explicit
' The macros behind the buttons of the Dashboard: read Jira, then build the flow sheets and the charts.
' Trivial entry points (like modButtons): the scenarios of flow_excel.feature run them against the fake Jira.
' A button macro must be a Public Sub without arguments.
'
' The button works in slices: BtnExtractAndCompute starts the extraction job (modExtract) and returns at once; one step
' of the job (a page of the search, a chunk of changelogs...) then runs per Application.OnTime call, so Excel stays usable,
' the bar of the Dashboard moves and the button Cancel works between two steps (and during a wait for Jira: modHttp reads
' gHttpCancel). Nothing of an earlier run is touched before the raw sheets are written (the last step of the extraction),
' so a cancel before that keeps the previous results.
' ExtractAndCompute is the same work in one go (no bar, no cancel): the tests and other macros use it.
' The workbook handler Workbook_BeforeClose (added by the build) calls AbortExtraction: a slice still scheduled when the
' workbook is closed would otherwise reopen it.

Private mBusy As Boolean
Private mCancel As Boolean
Private mExtracting As Boolean              ' True while the extraction job runs, False while the flow sheets are built
Private mNextSlice As Date

Private Function SliceName() As String
    SliceName = "'" & ThisWorkbook.Name & "'!ExtractSlice"       ' qualified: another open workbook must not answer the call
End Function

Public Sub ExtractAndCompute(Optional ByVal fromButton As Boolean = False)
    ' The optional argument keeps this Sub out of the Alt+F8 list: only the buttons (with their error handler) are for users.
    ' Everything comes from the workbook: Config (connection, proxy, filter, as-of date), Projects, Periods, Workflows.
    Dim num As Long, desc As String
    If mBusy Then Err.Raise vbObjectError + 611, "ExtractAndCompute", "An extraction is already running: wait for it or cancel it."
    On Error GoTo Failed
    Application.StatusBar = "Reading Jira..."
    ExtractFlowData                                ' empties the old results only once Jira has answered
    Application.StatusBar = "Building the flow sheets and the charts..."
    BuildFlowSheetsIn ThisWorkbook
    Application.StatusBar = False
    Exit Sub
Failed:
    num = Err.Number: desc = Err.Description
    Application.StatusBar = False
    Err.Raise num, "ExtractAndCompute", desc
End Sub

Public Sub BtnExtractAndCompute()
    If mBusy Then Beep: Exit Sub
    On Error GoTo Failed
    mBusy = True: mCancel = False: mExtracting = True: gHttpCancel = False
    ShowProgress 0, "Starting..."
    ExtractFlowDataStart
    ScheduleSlice
    Exit Sub
Failed:
    ExtractJobAbort
    EndRun "Stopped: " & Err.Number & " - " & Err.Description, 0
End Sub

Public Sub BtnCancelExtraction()
    ' Takes effect before the next step of the extraction, or at once during a wait for Jira; too late once the raw sheets
    ' are written.
    If Not mBusy Then Exit Sub
    If mExtracting Then
        mCancel = True
        gHttpCancel = True
        ShowProgress -1, "Cancelling..."
    Else
        ShowProgress -1, "Too late to cancel: the flow sheets are being built."
    End If
End Sub

Public Sub ExtractSlice()
    ' Called by Application.OnTime: one step, then the next call is scheduled. Keep it short (UI thread).
    If Not mBusy Then Exit Sub
    On Error GoTo Failed
    If mExtracting Then
        If mCancel Then
            ExtractJobAbort
            EndRun "Cancelled: the results of the previous run are kept.", 0
            Exit Sub
        End If
        If ExtractJobStep() Then
            mExtracting = False
            ShowProgress 0.92, "Building the flow sheets and the charts..."
        Else
            ShowProgress ExtractJobFraction(), ExtractJobLabel()
        End If
        ScheduleSlice
    Else
        BuildFlowSheetsIn ThisWorkbook
        EndRun "Done. The Dashboard shows the six flow metrics.", 1
    End If
    Exit Sub
Failed:
    If Err.Number = vbObjectError + 610 Then                         ' the Cancel button, pressed during a wait for Jira
        ExtractJobAbort
        EndRun "Cancelled: the results of the previous run are kept.", 0
        Exit Sub
    End If
    ExtractJobAbort
    EndRun "Stopped: " & Err.Number & " - " & Err.Description, 0
End Sub

Public Sub AbortExtraction()
    ' Forgets a pending slice (used when the workbook is closed or a test ends): an OnTime left behind would reopen the file.
    On Error Resume Next
    If mNextSlice <> 0 Then Application.OnTime mNextSlice, SliceName(), , False
    Err.Clear
    On Error GoTo 0
    mNextSlice = 0
    ExtractJobAbort
    mBusy = False: mCancel = False: mExtracting = False: gHttpCancel = False
End Sub

Public Function ProgressSnapshot() As String
    ' "fraction|status text" as the Dashboard shows them (one call for the tests that watch the bar from outside).
    Dim shapes As Object
    On Error Resume Next
    Set shapes = ThisWorkbook.Worksheets("Dashboard").Shapes
    ProgressSnapshot = Trim$(Str$(shapes("pbFill").Width / shapes("pbBack").Width)) & "|" & _
                       CStr(ThisWorkbook.Names("dashStatus").RefersToRange.Value)
End Function

Private Sub ScheduleSlice()
    mNextSlice = Now
    Application.OnTime mNextSlice, SliceName()
End Sub

Private Sub EndRun(ByVal message As String, ByVal fraction As Double)
    mBusy = False: mCancel = False: mNextSlice = 0: gHttpCancel = False
    Application.StatusBar = False
    ShowProgress fraction, message
End Sub

Private Sub ShowProgress(ByVal fraction As Double, ByVal text As String)
    ' The bar (rectangles pbBack and pbFill) and the status cell (name dashStatus) of the Dashboard; a fraction < 0 keeps the bar.
    Dim shapes As Object
    On Error Resume Next
    Set shapes = ThisWorkbook.Worksheets("Dashboard").Shapes
    If fraction >= 0 Then shapes("pbFill").Width = shapes("pbBack").Width * fraction + 0.1
    ThisWorkbook.Names("dashStatus").RefersToRange.Value = "'" & text
    Application.StatusBar = text
    If fraction >= 1 Or Not mBusy Then Application.StatusBar = False
    DoEvents
    Err.Clear
End Sub
