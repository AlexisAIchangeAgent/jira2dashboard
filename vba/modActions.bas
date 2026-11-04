Attribute VB_Name = "modActions"
Option Explicit
' The macro behind the button of the Dashboard: read Jira, then build the flow sheets and the charts.
' Trivial entry points (like modButtons): the scenario flow_excel.feature runs ExtractAndCompute against the fake Jira.
' A button macro must be a Public Sub without arguments. A busy flag stops a second click during a run.

Private mBusy As Boolean

Public Sub ExtractAndCompute(Optional ByVal fromButton As Boolean = False)
    ' The optional argument keeps this Sub out of the Alt+F8 list: only the button (with its error handler) is for users.
    ' Everything comes from the workbook: Config (connection, proxy, filter, as-of date), Projects, Periods, Workflows.
    Application.StatusBar = "Reading Jira..."
    ExtractFlowData                                ' empties the old results only once Jira has answered
    Application.StatusBar = "Building the flow sheets and the charts..."
    BuildFlowSheetsIn ThisWorkbook
    Application.StatusBar = False
End Sub

Public Sub BtnExtractAndCompute()
    If mBusy Then Beep: Exit Sub
    mBusy = True
    Application.Cursor = xlWait
    On Error GoTo Failed
    ExtractAndCompute
    Application.Cursor = xlDefault
    mBusy = False
    MsgBox "Done. The Dashboard shows the six flow metrics.", vbInformation, "jira2dashboard"
    Exit Sub
Failed:
    Application.Cursor = xlDefault
    Application.StatusBar = False
    mBusy = False
    MsgBox "Stopped: " & Err.Number & " - " & Err.Description, vbExclamation, "jira2dashboard"
End Sub
