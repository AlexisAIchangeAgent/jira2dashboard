Option Explicit
' Entry points for the buttons. A macro started by a button or by Alt+F8 must be a Public Sub
' without arguments: Functions and Subs with parameters do not appear in the Alt+F8 list.
' Keep these Subs trivial (one call, then a message). The automated tests never call them:
' a MsgBox in a hidden Excel would block the run.
' The runs use DoEvents, so Excel stays usable and buttons stay clickable: a busy flag blocks
' a second click (it would clear the results in the middle of a run).

Private mBusy As Boolean

Public Sub BtnRunAllTests()
    RunFromButton False
End Sub

Public Sub BtnRunNetworkTests()
    RunFromButton True
End Sub

Public Sub BtnClearResults()
    If mBusy Then Beep: Exit Sub
    ClearResults
End Sub

Private Sub ClearResults()
    ThisWorkbook.Worksheets("TestResults").Range("A:H").ClearContents
    ThisWorkbook.Worksheets("CoverageHits").Cells.ClearContents
    ProgressReset
End Sub

Private Sub RunFromButton(ByVal includeNetwork As Boolean)
    Dim summary As String
    If mBusy Then Beep: Exit Sub
    mBusy = True
    On Error GoTo Failed
    ClearResults
    Application.Cursor = xlWait
    summary = RunWithCoverage(includeNetwork)
    Application.Cursor = xlDefault
    Application.StatusBar = False
    Application.Calculate
    mBusy = False
    MsgBox "Tests: " & summary & vbCrLf & "Coverage: " & _
           Format$(ThisWorkbook.Worksheets("CoverageMap").Range("E2").Value2, "0.0%"), _
           vbInformation, "Test results"
    Exit Sub
Failed:
    Application.Cursor = xlDefault
    Application.StatusBar = False
    mBusy = False
    MsgBox "Tests stopped by an error: " & Err.Number & " - " & Err.Description, vbExclamation, "Test results"
End Sub
